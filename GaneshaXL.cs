using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelDna.Integration;
using ExcelDna.Registration;

// kudos to all that helped or inspired...
// https://docs.microsoft.com/en-us/office/client-developer/excel/calling-into-excel-from-the-dll-or-xll
// https://docs.microsoft.com/en-us/office/client-developer/excel/c-api-functions-that-can-be-called-only-from-a-dll-or-xll
// https://exceldna.typepad.com/blog/2006/01/some_new_exampl.html
// https://mikejuniperhill.blogspot.com/2014/12/writeread-dynamic-matrix-between-excel.html
// https://stackoverflow.com/questions/14896215/how-do-you-set-the-value-of-a-cell-using-excel-dna
// https://gist.github.com/govert/3012444
// https://groups.google.com/forum/#!topic/exceldna/MeYq0-LiGLM

namespace GaneshaXL
{
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            RegisterFunctions();
        }

        public void AutoClose()
        {
        }

        public void RegisterFunctions()
        {
            ExcelRegistration.GetExcelFunctions()
                             .Select(UpdateHelpTopic)
                             .RegisterFunctions();
        }

        public ExcelFunctionRegistration UpdateHelpTopic(ExcelFunctionRegistration funcReg)
        {
            funcReg.FunctionAttribute.HelpTopic = "http://www.duckduckgo.com";
            return funcReg;
        }
    }

    public class MyFunctions
    {
        private static Random rnd = new Random(Guid.NewGuid().GetHashCode());

        // cache parsed problem-schema used in gxRandom and gxParse
        private static FmtItem[] fmtArr = new FmtItem[1];
        private static FmtParser fmtParser = new FmtParser();
        private static FmtExtractor fmtExtractor = new FmtExtractor();
        private static uint fmtCRC = 0;
        private static Object fmtLock = new Object();
        private static StringBuilder fmtStringBuilder = new StringBuilder();

        private static void fmtCache(
            object[,] fmtObjArr
        )
        {
            int rows = fmtObjArr.GetLength(0);
            int cols = fmtObjArr.GetLength(1);

            lock (fmtLock)
            {
                // crc and cache a scan of the format row
                fmtStringBuilder.Clear();
                for (int col = 0; col < cols; col++)
                    fmtStringBuilder.Append((string)(fmtObjArr[0, col]));

                var crc = Crc32.Compute(Encoding.ASCII.GetBytes(fmtStringBuilder.ToString()));
                if (fmtCRC != crc)
                {
                    // reparse the format row as it has changed
                    fmtCRC = crc;
                    if (fmtArr.Length != cols)
                        fmtArr = FmtItem.CreateItemArr(cols);

                    fmtParser.SetCols(cols);
                    fmtParser.SetStateArr(fmtArr);
                    fmtExtractor.SetStateArr(fmtArr);

                    for (int col = 0; col < cols; col++)
                    {
                        var fmtStr = (string)(fmtObjArr[0, col]);
                        char[] fmtCharArr = fmtStr.ToCharArray();
                        Array.Reverse(fmtCharArr); // make operators postfix and numbers least-significant-first
                        fmtParser.Parse(fmtCharArr);
                    }
                }
            }
        }

        ///////////////////////////////////////////////////////////////

        /* Given nPk, determine the number of bits required for the compacted factoradic encoding
         * */
        [ExcelFunction(IsThreadSafe = true, Category = "GaneshaXL")]
        public static ushort gxPerm(ushort n, ushort k)
        {
            return Support.permutation_nbits(n, k);
        }

        /* Use the specified format row to determine the number of bits are needed for random 
         * population members then return randomized member data
         * */
        [ExcelFunction(IsThreadSafe = true, Category = "GaneshaXL")]
        public static object[,] gxRandom(
            [ExcelArgument(AllowReference = true)]object fmtRowArg,
            object rowsRowArg
        )
        {
            object[,] result = new object[1, 1];

            try
            {
                var fmtObjArr = (object[,])(((ExcelReference)fmtRowArg).GetValue());
                fmtCache(fmtObjArr);

                byte[] bytes0 = new byte[fmtParser.Bytesize()];

                int rows = Convert.ToInt32(rowsRowArg);
                result = new object[rows, 1];

                // return a column of randomized member data
                for (int row = 0; row < rows; row++)
                    result[row, 0] = System.Convert.ToBase64String(Support.RndFillbytes(bytes0));
            }
            catch
            {
                result[1,0] = ExcelError.ExcelErrorValue;
            }

            return result;
        }

        /* Decode the member data according to the columns in the specified format row
         * */
        [ExcelFunction(IsThreadSafe = true, Category = "GaneshaXL")]
        public static object[,] gxParse(
            [ExcelArgument(AllowReference = true)]object fmtRowArg,
            [ExcelArgument(AllowReference = true)]object dataColArg
        )
        {
            int row = 0, col = 0; // place outside to enhance error reporting
            int rows = 1, cols = 1;
            object[,] result = null;

            try
            {
                var fmtObjArr = (object[,])(((ExcelReference)fmtRowArg).GetValue());
                var dataObjArr = (object[,])(((ExcelReference)dataColArg).GetValue());

                fmtCache(fmtObjArr);

                rows = dataObjArr.GetLength(0);
                cols = fmtObjArr.GetLength(1);
                result = new object[rows, cols];

                // extract and cast the column fields from the member data, as formatted by the parser
                for (row = 0; row < rows; row++)
                {
                    fmtExtractor.SetRowData(System.Convert.FromBase64String((string)(dataObjArr[row, 0])));

                    for (col = 0; col < cols; col++)
                        result[row, col] = fmtExtractor.Extract(col);
                }
            }
            catch
            {
                // fill remainder with error
                if (result == null) result = new object[rows, cols];
                if (row == rows) row = 0; // reset?
                while (row < rows)
                {
                    for (col = 0; col < cols; col++)
                        result[row, col] = ExcelError.ExcelErrorValue;
                    row++;
                }
            }

            return result;
        }

        /* Using indexes to indicate which, cross two b64 encoded population members from a column
         * reference returning the b64 encoded result.
         * */
        [ExcelFunction(IsThreadSafe = true, Category = "GaneshaXL")]
        public static object[,] gxCross(
                [ExcelArgument(AllowReference = true)]object dataColArg,
                [ExcelArgument(AllowReference = true)]object idxColArg,
                [ExcelArgument(AllowReference = true)]object opt_mutprobArg
        )
        {
            int row = 0; // place outside loops to enhance error reporting
            int rows = 1;
            object[,] result = null;

            try
            {
                var dataObjArr = (object[,])(((ExcelReference)dataColArg).GetValue());
                var idxObjArr = (object[,])(((ExcelReference)idxColArg).GetValue());

                rows = idxObjArr.GetLength(0);
                result = new object[rows, 1];

                // assume a mutation rate, but change it if one is specified
                int mutprob = 10; // 1/10
                if (!(opt_mutprobArg is ExcelMissing))
                    switch (opt_mutprobArg)
                    {
                        case double d when opt_mutprobArg is double: // double was specified
                            mutprob = Math.Min(1, (int)(1 / d)); // turns 1/n into n
                            break;
                        case double d when ((ExcelReference)opt_mutprobArg).GetValue() is double: // ref to double was specified
                            mutprob = Math.Min(1, (int)(1 / d)); // turns 1/n into n
                            break;
                    }

                // return a cross (with mutation) of the specified population members
                // if indicies are the same, copy with no mutation
                for (row = 0; row < rows; row++)
                {
                    var i1 = Convert.ToUInt16((double)(idxObjArr[row, 0]));
                    var i2 = Convert.ToUInt16((double)(idxObjArr[row, 1]));
                    if (i1 == i2)
                        result[row, 0] = (string)(dataObjArr[i1, 0]);
                    else
                    {
                        byte[] barr1 = System.Convert.FromBase64String((string)(dataObjArr[i1, 0]));
                        byte[] barr2 = System.Convert.FromBase64String((string)(dataObjArr[i2, 0]));
                        string str = System.Convert.ToBase64String(Support.cross(barr1, barr2, mutprob));
                        result[row, 0] = (string)str;
                    }
                }
            }
            catch
            {
                // fill remainder with error
                if (result == null) result = new object[rows, 1];
                if (row == rows) row = 0; // reset?
                while(row < rows)
                    result[row++, 0] = ExcelError.ExcelErrorValue;
            }

            return result;
        }

        /* Calculate and return the most likely population indexes for crossover using the objective 
         * function values and the population member b64 data.
         * 
         * - duplicate population members are identified and filtered out
         * - the best performing member is preserved
         * - a sizable portion of the selections are made from the two best performing members
         * - a sizable portion of the selections are made from member's objective function expected value
         * - the remaining members are selected randomly
         * */
        [ExcelFunction(IsThreadSafe = true, Category = "GaneshaXL")]
        public static object[,] gxSample(
                [ExcelArgument(AllowReference = true)]object dataColArg,
                [ExcelArgument(AllowReference = true)]object objvalColArg
        )
        {
            int row = 0; // place outside to enhance error reporting
            int rows = 1;
            object[,] result = null;

            try
            {
                var dataObjArr = (object[,])(((ExcelReference)dataColArg).GetValue());
                var objvalObjArr = (object[,])(((ExcelReference)objvalColArg).GetValue());

                rows = objvalObjArr.GetLength(0);
                result = new object[rows, 2];

                var objvalArr = new float[rows];
                var expectedvalArr = new ushort[rows];

                // remove duplicates from the final selection by zeroing their objective value
                // determine duplicate from obj-value and data-crc
                var dupDict = new Dictionary<(uint, float), bool>();
                for (row = 0; row < rows; row++)
                {
                    // get abs of float-cast double objective-value
                    var objval = (float)Math.Abs((double)(objvalObjArr[row, 0]));

                    // get data-crc
                    var barr = System.Convert.FromBase64String((string)(dataObjArr[row, 0]));
                    var crcval = Crc32.Compute(barr);

                    if (dupDict.ContainsKey((crcval, objval)))
                        objval = 0; // duplicate
                    else
                        dupDict.Add((crcval, objval), true); // unique
                    objvalArr[row] = objval;
                }

                // identify the 2 best performing members
                float max1_val = 0, max2_val = 0, min_val = float.MaxValue;
                int max1_idx = 0, max2_idx = 0, min_idx = 0;
                for (row = 0; row < rows; row++)
                {
                    if (objvalArr[row] > max1_val)
                    {
                        max2_val = max1_val; max2_idx = max1_idx; // assign #1 to #2
                        max1_val = objvalArr[row]; max1_idx = row; // assign #1
                    }
                    else if (objvalArr[row] > max2_val)
                    {
                        max2_val = objvalArr[row]; max2_idx = row; // assign #2 only
                    }
                    if (objvalArr[row] < min_val)
                    {
                        min_val = objvalArr[row];
                        min_idx = row;
                    }
                }

                // if population is unhealthy, fill with random selections and exit
                float max_val = 0;
                max_val = Math.Max(0, Math.Max(max1_val, max2_val));
                if (max_val == 0)
                {
                    lock (rnd)
                    {
                        for (row = 0; row < rows; row++)
                        {
                            result[row, 0] = (ushort)rnd.Next(0, rows - 1);
                            result[row, 1] = (ushort)rnd.Next(0, rows - 1);
                        }
                    }
                    return result;
                }

                // apply nonlinear fitting to expected value, although a flatter scaling might lead to healthier populations
                // perhaps the fitting parameterized by iteration number? dunno...
                // apply 20x amplification to best member
                // determine the size of the distribution array
                const float amplifier = 20;
                int distribSize = 0;
                for (row = 0; row < rows; row++)
                {
                    expectedvalArr[row] = (ushort)Math.Round(amplifier * Math.Pow(objvalArr[row] / max_val, 2));
                    distribSize += expectedvalArr[row];
                }

                // allocate and fill the distribution array with the expected members
                var distribArr = new ushort[distribSize];
                {
                    int i = 0;
                    for (row = 0; row < rows; row++)
                        for (int ev = expectedvalArr[row]; ev > 0; ev--)
                            distribArr[i++] = (ushort)row;
                }

                //////////////////////////////////
                // configure returned selection 

                // initially fill with random selections from the distribution
                lock (rnd)
                {
                    for (row = 0; row < rows; row++)
                    {
                        result[row, 0] = distribArr[(ushort)rnd.Next(0, distribSize - 1)];
                        result[row, 1] = distribArr[(ushort)rnd.Next(0, distribSize - 1)];
                    }
                }

                // then, ensure the best two members get a larger amount of breeding
                // by forcing column 0 to be their index.
                {
                    row = 0;
                    for (; row < (int)(rows * 1 / 6); row++) result[row, 0] = max1_idx;
                    for (; row < (int)(rows * 2 / 6); row++) result[row, 0] = max2_idx;
                }

                // finally, ensure both max1_idx and max2_idx live on via self-breeding
                result[0, 1] = result[0, 0];
                result[(int)(rows * 1 / 6), 1] = result[(int)(rows * 1 / 6), 0];
            }
            catch
            {
                // fill remainder with error
                if (result == null) result = new object[rows, 2];
                if (row == rows) row = 0; // reset?
                while (row < rows)
                {
                    result[row, 0] = ExcelError.ExcelErrorValue;
                    result[row, 1] = ExcelError.ExcelErrorValue;
                    row++;
                }
            }

            return result;
        }

    }
}
