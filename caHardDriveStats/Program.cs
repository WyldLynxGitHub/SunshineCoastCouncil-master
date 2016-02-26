using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace caHardDriveStats
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("File size: " + ((GetDirectorySize(@"D:\SCC\ECM\")/1024)/1024).ToString());
        }
        public static long GetDirectorySize(string parentDirectory)
        {
            return new DirectoryInfo(parentDirectory).GetFiles("*.*", SearchOption.AllDirectories).Sum(file => file.Length);
        }
    }
}
