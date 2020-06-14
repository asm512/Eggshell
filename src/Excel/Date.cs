using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel
{
    public class Date : IExcelUtility
    {
        private const string ResultPath = "_DATE_EXPORT.txt";

        public Date()
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
                File.AppendAllText(ResultPath, formattedDate);
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
                File.AppendAllText(ResultPath, formattedDate);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                Console.ReadKey();
            }
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
