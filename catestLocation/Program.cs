using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace catestLocation
{
    class Program
    {
        static void Main(string[] args)
        {
            string dfd = @"\\scas57\ECMIncache18\";
            if (dfd.Length == 21)
            {
                Console.WriteLine("Single ext: " + dfd.Substring(9, 11));
            }
            else if(dfd.Length == 22)
            {
                Console.WriteLine("Single ext: " + dfd.Substring(9, 12));
            }


            //Console.WriteLine("Double: " + dfd.Length.ToString());
            //Console.WriteLine("Double ext: " + dfd.Substring(9, 12));
            //string dfdf = @"\\scas57\ECMIncache4\";
            //Console.WriteLine("Single: " + dfdf.Length.ToString());
            
        }
    }
}
