using System;

namespace Config
{
    internal class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Input dirPath:");
                var dirPath = Console.ReadLine();
                Console.WriteLine("Input outPath:");
                var outPath = Console.ReadLine();
                ExcelReaderSystem.Program(dirPath, outPath);
                Console.WriteLine("Done! Press any key to exit...");
                return;
            }

            if (args.Length >= 3 && args[0] == "-path")
            {
                ExcelReaderSystem.Program(args[1], args[2]);
            }
        }
    }
}