using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HP.HPTRIM.SDK;

namespace caTestCallingOrigins
{
    class Program
    {
        static void Main(string[] args)
        {
            TrimApplication.Initialize();
            using (Database db = new Database())
            {
                OriginHistory  fdfd = (OriginHistory)db.FindTrimObjectByUri(BaseObjectTypes.OriginHistory, 119);
                if(fdfd!=null)
                {
                    var lst = fdfd.GetChildObjectList(BaseObjectTypes.Record);
                    foreach(Record r in lst)
                    {
                        Console.WriteLine(r.Number + Environment.NewLine);
                    }
                }
            }
        }
    }
}
