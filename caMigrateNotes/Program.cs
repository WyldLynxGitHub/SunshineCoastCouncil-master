using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HP.HPTRIM.SDK;

namespace caMigrateNotes
{
    class Program
    {
        static void Main(string[] args)
        {
            TrimApplication.Initialize();
            using (Database db = new Database())
            {
                TrimMainObjectSearch objS = new TrimMainObjectSearch(db, BaseObjectTypes.Record);
                objS.SetSearchString("");

            }
        }
    }
}
