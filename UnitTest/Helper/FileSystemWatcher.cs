using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting.Logging;

namespace PP5AutoUITests
{
    public class FileSystemWatcherWrapper
    {
        // 定义一个事件，供主程序订阅
        public event FileSystemEventHandler FileChanged;
        public List<FileSystemWatcher> watchers = new List<FileSystemWatcher>();
        //public string[] filters;

        public FileSystemWatcherWrapper() {}

        public void CreateFileWatcher_FullFileName(string path, NotifyFilters notifyFilters, string fileToWatch)
        {
            // Create a new FileSystemWatcher and set its properties.
            FileSystemWatcher watcher = new FileSystemWatcher();
            watcher.Path = path;
            /* Watch for changes in CreationTime and LastWrite */
            watcher.NotifyFilter = notifyFilters;
            // Watch files with given extension, default is txt
            watcher.Filter = fileToWatch;
            //watcher.Filter = "";

            // Add event handlers.
            watcher.Changed += new FileSystemEventHandler(OnChanged);
            watcher.Created += new FileSystemEventHandler(OnChanged);
            watcher.Deleted += new FileSystemEventHandler(OnChanged);
            watcher.Renamed += new RenamedEventHandler(OnRenamed);

            // Begin watching.
            watcher.EnableRaisingEvents = true;
            //watcher.SynchronizingObject = this;
            watchers.Add(watcher);
        }

        public void CreateFileWatcher(string path, NotifyFilters notifyFilters, string fileExtToWatch = "txt")
        {
            // Create a new FileSystemWatcher and set its properties.
            FileSystemWatcher watcher = new FileSystemWatcher();
            watcher.Path = path;
            /* Watch for changes in CreationTime and LastWrite */
            watcher.NotifyFilter = notifyFilters;
            // Watch files with given extension, default is txt
            watcher.Filter = $"*.{fileExtToWatch}";
            //watcher.Filter = "";

            // Add event handlers.
            watcher.Changed += new FileSystemEventHandler(OnChanged);
            watcher.Created += new FileSystemEventHandler(OnChanged);
            watcher.Deleted += new FileSystemEventHandler(OnChanged);
            watcher.Renamed += new RenamedEventHandler(OnRenamed);

            // Begin watching.
            watcher.EnableRaisingEvents = true;
            //watcher.SynchronizingObject = this;
            watchers.Add(watcher);
        }

        public void CreateFileWatchers(string[] paths, NotifyFilters notifyFilters, string[] fileExtsToWatch)
        {
            for(int i = 0; i < paths.Length; i++)
            {
                CreateFileWatcher(paths[i], notifyFilters, fileExtsToWatch[i]);
            }
        }

        // Define the event handlers.
        private void OnChanged(object source, FileSystemEventArgs e)
        {
            // Specify what is done when a file is changed, created, or deleted.
            Logger.LogMessage("File: " + e.FullPath + " " + e.ChangeType);

            // get the file's extension 
            string strFileExt = Path.GetExtension(e.FullPath);

            // filter file types 
            //if (filters.Contains(strFileExt))
            //{
                //Console.WriteLine("watched file type changed.");
                if (ThreadHelper.WaitUntil(() => FileIsReady(e.FullPath), 500, 10))
                {
                    // 触发外部事件
                    FileChanged?.Invoke(this, e);          // Check files in the Main thread otherwise threading issues occur
                                                           //FileChanged?.BeginInvoke((MethodInvoker)(() => SomeMethod()));  
                }
            //}
        }

        private void OnRenamed(object source, RenamedEventArgs e)
        {
            // Specify what is done when a file is renamed.
            Logger.LogMessage("File: {0} renamed to {1}", e.OldFullPath, e.FullPath);
        }

        private bool FileIsReady(string path)
        {
            //One exception per file rather than several like in the polling pattern
            try
            {
                //If we can't open the file, it's still copying
                using (var file = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    return file.Length > 0;
                }
            }
            catch (IOException)
            {
                return false;
            }
        }

        public void Dispose()
        {
            foreach (var watcher in watchers)
                watcher?.Dispose();
        }
    }
}
