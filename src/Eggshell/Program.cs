using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;

namespace Eggshell
{
    class Program
    {
        public struct ExcelLibInstances
        {
            public static Excel.DateFormat _dateLib = new Excel.DateFormat();
        }

        static void Main(string[] args)
        {
            Console.Title = "Eggshell Utility";

            if (args.Length != 1)
            {
                Console.WriteLine("Incorrect args length");
                Console.ReadKey();
                return;
            }

            Console.WriteLine($"Using -> {args[0]}{Environment.NewLine}");
            DISPLAY_FUNCTIONS();
            PARSE_INPUT_DEBUG(args);
            Console.WriteLine("Operation finished");
            Console.ReadKey();
        }

        public static void DISPLAY_FUNCTIONS()
        {
            Console.WriteLine("CHANGE INDEX OF D M Y VALUES");
            Console.WriteLine("/idx d.d.m.m.y.y.y.y    (8) where array references indexes ddmmyyyy" + Environment.NewLine);

            Console.WriteLine("CONVERT DATES CONTAINING JAN,FEB,MARCH ETC");
            Console.WriteLine("/fmt d.d.m.m.m.y.y.y.y   (9) where array references indexes ddmmmyyyy" + Environment.NewLine);

            Console.WriteLine("CONVERT QUARTERLY INTERVALS TO MONTHLY INTERVALS ddmmyyyy");
            Console.WriteLine("/q2m" + Environment.NewLine);
        }

        public static void PARSE_INPUT(string[] args)
        {
            try
            {
                string cmd = Console.ReadLine().ToLower();

                switch ($"{cmd[0]}{cmd[1]}{cmd[2]}{cmd[3]}")
                {
                    case "/idx":
                        foreach (var toProc in File.ReadAllLines(args[0]))
                        {
                            ExcelLibInstances._dateLib.CONVERT_STRING_TO_DATE_INDEX_CHANGE(toProc, STRING_TO_INT_VECTOR(cmd.Replace("/idx ", "")));
                        }
                        return;
                    case "/fmt":
                        foreach (var toProc in File.ReadAllLines(args[0]))
                        {
                            ExcelLibInstances._dateLib.CONVERT_STRING_TO_DATE_FORMAT_CHANGE(toProc, STRING_TO_INT_VECTOR(cmd.Replace("/fmt ", "")));
                        }
                        return;
                    case "/q2m":
                        ExcelLibInstances._dateLib.CONVERT_QUARTERLY_TO_MONTHLY(File.ReadAllLines(args[0]));
                        return;
                    default:
                        Console.WriteLine("Command not recognised");
                        PARSE_INPUT(args);
                        return;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return;
            }

        }

        public static void PARSE_INPUT_DEBUG(string[] args)
        {
            string cmd = Console.ReadLine().ToLower();

            switch ($"{cmd[0]}{cmd[1]}{cmd[2]}{cmd[3]}")
            {
                case "/idx":
                    foreach (var toProc in File.ReadAllLines(args[0]))
                    {
                        ExcelLibInstances._dateLib.CONVERT_STRING_TO_DATE_INDEX_CHANGE(toProc, STRING_TO_INT_VECTOR(cmd.Replace("/idx ", "")));
                    }
                    return;
                case "/fmt":
                    foreach (var toProc in File.ReadAllLines(args[0]))
                    {
                        ExcelLibInstances._dateLib.CONVERT_STRING_TO_DATE_FORMAT_CHANGE(toProc, STRING_TO_INT_VECTOR(cmd.Replace("/fmt ", "")));
                    }
                    return;
                case "/q2m":
                    ExcelLibInstances._dateLib.CONVERT_QUARTERLY_TO_MONTHLY(File.ReadAllLines(args[0]));
                    return;
                default:
                    Console.WriteLine("Command not recognised");
                    PARSE_INPUT(args);
                    return;
            }
        }

        private static int[] STRING_TO_INT_VECTOR(string s)
        {
            string[] intArrasStr = s.Split('.');
            var r = new List<int>();
            foreach (var c in intArrasStr)
            {
                r.Add(Convert.ToInt16(c));
            }
            return r.ToArray();
        }
    }
}
