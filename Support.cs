using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using ExcelDna.Registration;

// https://docs.microsoft.com/en-us/office/client-developer/excel/calling-into-excel-from-the-dll-or-xll
// https://docs.microsoft.com/en-us/office/client-developer/excel/c-api-functions-that-can-be-called-only-from-a-dll-or-xll
// https://exceldna.typepad.com/blog/2006/01/some_new_exampl.html
// https://mikejuniperhill.blogspot.com/2014/12/writeread-dynamic-matrix-between-excel.html
// https://stackoverflow.com/questions/14896215/how-do-you-set-the-value-of-a-cell-using-excel-dna
// https://gist.github.com/govert/3012444

namespace GaneshaXL
{
    public class NumParser
    {
        public double fracpart = 0;
        public int signpart = 1;
        public uint intpart = 0;
        public uint exponent = 1;

        public NumParser() { Clear(); }
        public void Clear() { signpart = 1; fracpart = 0; intpart = 0; exponent = 1; }
        public void Negate() { signpart = -1; }
        public void ConvertFrac()
        {
            fracpart = 0;
            while (exponent >= 10)
            {
                fracpart += (double)(intpart % 10) / (double)exponent;
                exponent /= 10;
                intpart /= 10;
            }
            intpart = 0; exponent = 1;
        }
        public void Extend(char ch)
        {
            intpart += (uint)Char.GetNumericValue(ch) * exponent;
            exponent *= 10;
        }
        public double ExtractFloat() { return (double)signpart * ((double)intpart + fracpart); }
        public ushort ExtractUInt() { return Convert.ToUInt16(intpart); }
    };

    public class BitExtractor
    {
        int bit;
        byte[] bytes; // ref?
        public BitExtractor(byte[] b) { bytes = b; bit = 0; }
        public ulong Extract(ushort nbits)
        {
            if (nbits > 64) return 0;

            ulong r = 0;
            for (int i = 0; i < nbits && bit < 8 * bytes.Length; i++)
            {
                var by = bytes[bit / 8]; // left chars first
                var bs = 7 - (bit % 8); // high bits first
                var bm = 1 << bs;
                if ((by & bm) != 0) r |= (ulong)1 << ((int)nbits - i - 1); // high bits first
                bit += 1;
            }
            return r;
        }
    }

    // https://en.wikipedia.org/wiki/Lehmer_code
    // https://en.wikipedia.org/wiki/Factorial_number_system
    // 64bit order can only enumerate a max of 20 resources
    public class Perm64
    {
        private ulong s; // serial number
        private ushort n, k; // nPk
        private ushort[] remainders; // state
        private ushort[] indicies; // state (circular)
        public ushort[] symbols; // output

        public Perm64(ulong s0, ushort n0, ushort k0)
        {
            s = s0; n = n0; k = k0;
            symbols = new ushort[k]; // symbols arr [0..k)
            remainders = new ushort[k + 1]; // remainders [0..k] // ]? // +1?
            indicies = new ushort[n];

            // performs a variant of the factoradic coding
            for (ushort i = 0; i < n; i++) indicies[i] = i;

            try
            {
                // calc remainders
                {
                    ushort k1 = k, n1 = n;
                    ulong s1 = s;                       // make working copies
                    int r = 0;                          // remainder #
                    do
                    {
                        remainders[r++] = (ushort)(s1 % n1);    // store remainder
                        s1 /= n1--;                             // carry forward quotient and reduce radix
                    } while (--k1 > 0);                         // while there are more selections to make

                    while (r < k + 1) remainders[r++] = 0;   // clear remaining remainders...
                }

                // extract symbols
                for (ushort i = 0, front = 0; i < k; i++)
                {
                    ulong r = remainders[i];
                    symbols[i] = indicies[(ushort)(front + r) % n];
                    indicies[(ushort)(front + r) % n] = indicies[front];
                    front = (ushort)((front + 1) % n);
                }
            }
            catch
            {
                // store some obviously weirdly large number...
                for (ushort i = 0; i < k; i++) symbols[i] = 0xFFFF;
                throw new Exception(); // rethrow
            }
        }
    };

    /* Specifies how a field-item should be formatted
     * */
    public class FmtItem
    {
        public char fmt;
        public ushort bit_width;
        public ulong bit_mask;
        //
        public double float_offset, float_step;
        public ushort float_count;
        //
        public ushort perm_n, perm_k, perm_group, perm_pick;
        //
        public FmtItem(char ch) { fmt = ch; }
        public static FmtItem[] CreateItemArr(int cols)
        {
            FmtItem[] states = new FmtItem[cols];
            for (int i = 0; i < cols; i++) states[i] = new FmtItem('e');
            return states;
        }
    };

    /* Parses the format row that specifies how field-items should be extracted from raw member data
     * */
    public class FmtParser
    {
        public int cols, col;
        private FmtItem[] states = null;
        const string numbers = "0123456789";
        NumParser num = new NumParser();
        public Stack<double> stack = new Stack<double>();

        public void SetCols(int _cols)
        {
            cols = _cols;
            col = 0;
        }

        public void SetStateArr(FmtItem[] s)
        {
            states = s;
        }

        public void PushNum() { stack.Push(num.ExtractFloat()); }
        public double PopFloat() { return stack.Pop(); }
        public ushort PopUInt() { return Convert.ToUInt16(Math.Abs(stack.Pop())); }
        public ushort Bitsize()
        {
            ushort n = 0;
            for (int i = 0; i < cols; i++)
                if (states[i].fmt == 'p')
                    n += (states[i].perm_pick == 0) ? states[i].bit_width : (ushort)0; // only add bits once per group
                else
                    n += states[i].bit_width;
            return n;
        }

        public ushort Bytesize()
        {
            var n = (ushort)((1 + Bitsize()) >> 3);
            return (n == (ushort)0) ? (ushort)1 : n;
        }
        public void Parse(char[] fmt)
        {
            num.Clear();
            stack.Clear();
            try
            {
                foreach (char ch in fmt)
                {
                    switch (ch)
                    {
                        /////////////// basic types
                        case char n when numbers.Contains(ch):
                            num.Extend(n);
                            break;
                        case '.': // set fracpart
                            num.ConvertFrac();
                            break;
                        case '-': // negate
                            num.Negate();
                            break;
                        case ',': // save arg
                            PushNum();
                            num.Clear();
                            break;
                        /////////////// higher level types
                        case 'n': // nil field
                            states[col].fmt = 'n';
                            states[col].bit_width = 0;
                            states[col].bit_mask = 0;
                            num.Clear();
                            break;
                        case 'b': // bits as integer: b,<bitlength>
                            states[col].fmt = (stack.Count == 0) ? 'b' : 'e';
                            states[col].bit_width = num.ExtractUInt();
                            states[col].bit_mask = (uint)(1 << (int)states[col].bit_width) - 1;
                            num.Clear();
                            break;
                        case 'f': // float: f,<offset>,<stepsize>,<stepcount>
                            states[col].fmt = (stack.Count == 2) ? 'f' : 'e';
                            states[col].float_offset = num.ExtractFloat(); // arg1
                            states[col].float_step = PopFloat(); // arg2
                            states[col].float_count = PopUInt(); // arg3
                            states[col].bit_width = Support.count_nbits(states[col].float_count);
                            states[col].bit_mask = (ulong)(1 << (int)states[col].bit_width) - 1;
                            num.Clear();
                            break;
                        case 'p': // permutation: p,<n>,<k>,<group>,<pick>
                            states[col].fmt = (stack.Count == 3) ? 'p' : 'e';
                            states[col].perm_n = num.ExtractUInt(); // arg1
                            states[col].perm_k = PopUInt(); // arg2;
                            states[col].perm_group = PopUInt(); // arg3;
                            states[col].perm_pick = PopUInt(); // arg4;
                            states[col].bit_width = Support.permutation_nbits(states[col].perm_n, states[col].perm_k);
                            states[col].bit_mask = (ulong)(1 << (int)states[col].bit_width) - 1;
                            num.Clear();
                            break;
                    }
                }
            }
            catch
            {
                states[col].fmt = 'e'; // err
            }
            col++;
        }
    }

    public class FmtExtractor
    {
        public FmtItem[] states = null;
        public Perm64[] permArr = null;
        public BitExtractor bitextractor;

        public void SetStateArr(FmtItem[] s)
        {
            states = s;

            int n = 0;
            for (int i = 0; i < states.Length; i++)
                n = Math.Max(n, states[i].perm_group);
            permArr = new Perm64[n +1];
        }

        public void SetRowData(byte[] barr)
        {
            for (int i = 0; i < permArr.Length; i++)
                permArr[i] = null;

            bitextractor = new BitExtractor(barr);
        }

        public object Extract(int col)
        {
            switch (states[col].fmt)
            {
                case 'n': // nil field
                    return ExcelError.ExcelErrorNA;
                case 'e': // force an error
                    return ExcelError.ExcelErrorValue;
                case 'b': // extract raw bits
                    return bitextractor.Extract(states[col].bit_width);
                case 'f': // extract a float
                    return (double)((int)bitextractor.Extract(states[col].bit_width) % (int)states[col].float_count) * states[col].float_step + states[col].float_offset;
                case 'p': // extract permutation data, for specified perm group
                    if (permArr[states[col].perm_group] == null)
                        permArr[states[col].perm_group] = new Perm64(bitextractor.Extract(states[col].bit_width), states[col].perm_n, states[col].perm_k);
                    if (states[col].perm_pick > permArr[states[col].perm_group].symbols.Length - 1)
                        return ExcelError.ExcelErrorValue; // bad perm group
                    else
                        return (int)permArr[states[col].perm_group].symbols[states[col].perm_pick];
                default:
                    return ExcelError.ExcelErrorValue;
            }
        }
    }

    public class Support
    {
        private static Random rnd = new Random(Guid.NewGuid().GetHashCode());

        public static byte[] RndFillbytes(byte[] bytes)
        {
            lock (rnd)
            {
                rnd.NextBytes(bytes);
            }
            return bytes;
        }

        /* GA cross algorithm that performs a bitwise splice between two members followed by one mutation of the result
         * */
        public static byte[] cross(byte[] b0, byte[] b1, int mutprob)
        {
            if (b0.Length != b1.Length)
                throw new Exception();

            byte[] b2 = new byte[b0.Length];

            int swap, splicebit, mut, mutbit, bytei, biti;
            lock (rnd)
            {
                swap = rnd.Next();
                mut = rnd.Next(1, mutprob);
                splicebit = rnd.Next(b0.Length * 8);
                mutbit = rnd.Next(b0.Length * 8);
            }
            bytei = splicebit >> 3;
            biti = splicebit & 0x07;
            byte mask = (byte)((1 << biti) - 1);

            if ((swap & 1) == 1)
            {
                for (int i = 0; i < bytei; i++) b2[i] = b0[i]; // starts b0
                b2[bytei] = (byte)((b0[bytei] & mask) | (b1[bytei] & ~mask));
                bytei++;
                for (; bytei < b0.Length; bytei++) b2[bytei] = b1[bytei]; // ends b1
            }
            else
            {
                for (int i = 0; i < bytei; i++) b2[i] = b1[i]; // starts b1
                b2[bytei] = (byte)((b1[bytei] & mask) | (b2[bytei] & ~mask));
                bytei++;
                for (; bytei < b0.Length; bytei++) b2[bytei] = b0[bytei]; // ends b0
            }
            if (mut == 1) b2[mutbit >> 3] ^= (byte)(1 << (mutbit & 0x07)); // mutate
            return b2;
        }

        /* for our purposes, Ramanujan_ln is a better Gamma_ln function than Sterling_ln because it
         * slightly overestimates, rather underestimates bitlength (domainspace)
         * */
        public static double ramanujan_ln(double n)
        {
            return n * Math.Log(n) - n + Math.Log(n * (1 + 4 * n * (1 + 2 * n)) + 1 / 30) / 6 + .5 * Math.Log(Math.PI);
        }

        public static ushort permutation_nbits(ushort n, ushort k)
        {
            if (n == 0 || k == 0) return 0;

            if (n == k) return Convert.ToUInt16(ramanujan_ln((double)n) / Math.Log(2) + 1); // plain ln2 of Gamma

            double permut_ln = ramanujan_ln((double)n) - ramanujan_ln((double)(n - k));
            return Convert.ToUInt16(permut_ln / Math.Log(2) + 1); // ln2 of nPk
        }

        /* Given n, determine the number of bits required to represent it 
         * */
        public static ushort count_nbits(ushort count)
        {
            return Convert.ToUInt16(Math.Round(Math.Log(count, 2)));
        }

    }
}
