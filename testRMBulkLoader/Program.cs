using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HP.HPTRIM.SDK;


namespace testRMBulkLoader
{
    class Program
    {
        static void Main(string[] args)
        {
            TrimApplication.Initialize();
            using(Database db = new Database())
            {
                //db.Id="45";
                db.Connect();
            bulkLoaderSample bls = new bulkLoaderSample();
            bls.run(db);
            }
        }
    }
}
