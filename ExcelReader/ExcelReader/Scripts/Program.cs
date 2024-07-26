using System;

namespace Config
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string dirPath;
            string outPath;

            if (args.Length == 0)
            {
                Console.WriteLine("Input dirPath:");
                dirPath = Console.ReadLine();
                Console.WriteLine("Input outPath:");
                outPath = Console.ReadLine();
                ExcelReaderSystem.Program(dirPath, outPath);
                Console.WriteLine("Done! Press any key to exit...");
                return;
            }

            if (args.Length >= 3 && args[0] == "-path")
            {
                dirPath = args[1];
                outPath = args[2];
            }
            else
            {
                return;
            }

            var isFormat = args.Length >= 4 && args[3] == "-format";
            ExcelReaderSystem.Program(dirPath, outPath, isFormat);
        }
    }
}