using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel
{
    public class DateFormat : IExcelUtility
    {
        private const string ResultPath = "_DATE_EXPORT.txt";

        public DateFormat()
        {
            if (File.Exists(ResultPath))
            {
                File.Delete(ResultPath);
            }
        }

        public string GET_RESULT_PATH()
        {
            return ResultPath;
        }

        public void CONVERT_STRING_TO_DATE_INDEX_CHANGE(string toParse, int[] indexArr)
        {
			try
			{
                string formattedDate = $"{toParse[indexArr[0]]}{toParse[indexArr[1]]}/{toParse[indexArr[2]]}{toParse[indexArr[3]]}/{toParse[indexArr[4]]}{toParse[indexArr[5]]}{toParse[indexArr[6]]}{toParse[indexArr[7]]}{Environment.NewLine}";
                File.AppendAllText(GET_RESULT_PATH(), formattedDate);
			}
			catch (Exception e)
			{
                Console.WriteLine(e.ToString());
                Console.ReadKey();
            }
        }

        public void CONVERT_STRING_TO_DATE_FORMAT_CHANGE(string toParse, int[] indexArr)
        {
            if (indexArr.Length != 9) { Console.WriteLine("Incorrect arg length"); return; }
            try
            {
                string formattedDate = $"{toParse[indexArr[0]]}{toParse[indexArr[1]]}/";
                formattedDate += MONTH_TO_INT($"{toParse[indexArr[2]]}{toParse[indexArr[3]]}{toParse[indexArr[4]]}").ToString();
                formattedDate += $"/{toParse[indexArr[5]]}{toParse[indexArr[6]]}{toParse[indexArr[7]]}{toParse[indexArr[8]]}{Environment.NewLine}";
                File.AppendAllText(GET_RESULT_PATH(), formattedDate);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                Console.ReadKey();
            }
        }

        public void CONVERT_QUARTERLY_TO_MONTHLY(string[] toParse)
        {
            var _sb = new StringBuilder();

            foreach (var q in toParse)
            {
                //Add the start of the quarter
                _sb.Append(q + Environment.NewLine);

                //Add the next 2 months
                int quarterStart = Convert.ToInt32($"{q[3]}{q[4]}");

                var qm1 = q[0].ToString() + q[1].ToString() + q[2].ToString() + (quarterStart + 1).ToString() + q[5].ToString() + q[6].ToString() + q[7].ToString() + q[8].ToString() + q[9].ToString() + Environment.NewLine;
                _sb.Append(qm1);

                var qm2 = q[0].ToString() + q[1].ToString() + q[2].ToString() + (quarterStart + 2).ToString() + q[5].ToString() + q[6].ToString() + q[7].ToString() + q[8].ToString() + q[9].ToString() + Environment.NewLine;
                _sb.Append(qm2);
            }

            File.WriteAllText(GET_RESULT_PATH(), _sb.ToString());
        }

        public enum Month
        {
            Jan = 01,
            Feb = 02,
            Mar = 03,
            Apr = 04,
            May = 05,
            Jun = 06,
            Jul = 07,
            Aug = 08,
            Sep = 09,
            Oct = 10,
            Nov = 11,
            Dec = 12,
        }

        public static int MONTH_TO_INT(string month)
        {
            switch (month.ToLower())
            {
                case "jan":
                    return (int)Month.Jan;
                case "feb":
                    return (int)Month.Feb;
                case "mar":
                    return (int)Month.Mar;
                case "apr":
                    return (int)Month.Apr;
                case "may":
                    return (int)Month.May;
                case "jun":
                    return (int)Month.Jun;
                case "jul":
                    return (int)Month.Jul;
                case "aug":
                    return (int)Month.Aug;
                case "sep":
                    return (int)Month.Sep;
                case "oct":
                    return (int)Month.Oct;
                case "nov":
                    return (int)Month.Nov;
                case "dec":
                    return (int)Month.Dec;
                default:
                    Console.WriteLine($"could not process val {month}");
                    return 0;
            }
        }
    }
}
