using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FolderWatcher
{
    
    class Program
    {
        
        static void Main(string[] args)
        {
            var fw = new FileSystemWatcher(@"C:\temp\T1Watcher\", "*.txt");
            while (true)
            {
                Console.WriteLine("added file {0}",
                    fw.WaitForChanged(WatcherChangeTypes.All).Name);
            }

            //System.IO.FileSystemWatcher watcher;
            //watcher = new System.IO.FileSystemWatcher();
            //watcher.Path = @"C:\temp\T1Watcher\";
            //watcher.EnableRaisingEvents = true;
            //watcher.Filter = "*.txt";
            //watcher.NotifyFilter = System.IO.NotifyFilters.LastWrite;
            //watcher.Changed += new System.IO.FileSystemEventHandler(FileChanged);
        }

        //private static void FileChanged(object sender, System.IO.FileSystemEventArgs e)
        //{
        //    if (!IsFileReady(@"C:\temp\T1Watcher\")) return;
        //    //DoWorkOnFile(e.FullPath);
        //}

        //private static bool IsFileReady(string path)
        //{
        //    try
        //    {
        //        //If we can't open the file, it's still copying
        //        using (var file = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.Read))
        //        {
        //            return true;
        //        }
        //    }
        //    catch (IOException)
        //    {
        //        return false;
        //    }

        //}
      

    }
}
