using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HP.HPTRIM.SDK;

namespace TestListMax
{
    class Program
    {
        static void Main(string[] args)
        {
            TrimApplication.Initialize();
            using(Database db = new Database())
            {
                for (int i = 0; i < 94; i++)
                {
                    if (i == 98)
                    {
                        var sdsds = "ewrwer";
                    }
                    //Classification cl = new Classification(db, 9144);
                    Classification par = new Classification(db, 9144);
                    Classification c = new Classification(db);
                    c = par.NewLowerLevel();
                    c.Name = "EPL Test " + Guid.NewGuid().ToString();

                        c.RecordPattern = "/gggg";

                    c.LevelNumber = "/" + nextnumber().ToString().PadLeft(4, '0');
                    c.Save();
                }


            }
            //List<int> test = new List<int>();

            //for (int i = 1; i < 105; i++)
            //{
            //    int ibla = i * 10;
            //    test.Add(ibla);
            //    Console.WriteLine("Max test: " + test.Max().ToString() + " Last test: " + test.Last().ToString());
            //}
            
        }
        private static int nextnumber()
        {
            TrimApplication.Initialize();
            using (Database db = new Database())
            {
                int intChildCnt = 0;
                int intChildLast = 0;
                List<int> classNumb = new List<int>();
                List<Classification> cl = new List<Classification>();

                TrimMainObjectSearch obj = new TrimMainObjectSearch(db, BaseObjectTypes.Classification);

                obj.SetSearchString("parent:" + 9144);
                obj.SetSortString("numberx");
                int icnt = 1;
                if (obj.FastCount > 0)
                {
                    foreach (Classification c in obj)
                    {
                        if (icnt > 98)
                        {
                            var sdsd = icnt;
                        }
                        cl.Add(c);
                        var sdsdsd = c.LevelNumberUncompressed;
                        var kkkk = c.LevelNumberUncompressed.Remove(0, 1);
                        classNumb.Add(Convert.ToInt32(kkkk));

                        icnt++;
                    }
                    //classNumb.Sort(x=>x)
                    //cl.Max(x => x.Uri);
                    intChildCnt = classNumb.Max();
                    intChildLast = classNumb.Last();
                }
                else
                {
                    intChildCnt = 0;
                }
                Console.WriteLine(" Container count: " + obj.FastCount.ToString() + " Max number: " + (intChildCnt + 10).ToString() + " Last number: " + (intChildLast + 10).ToString());
                return intChildLast + 10;
            }
        }
    }
}
