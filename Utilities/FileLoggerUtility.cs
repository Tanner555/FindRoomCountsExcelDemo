using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace MyCommonUtilities
{
    public class FileLoggerUtility
    {
        bool bLoggedFirstMessage = false;
        public FileInfo _fileInfo;

        /// <summary>
        /// Creates New Instance Of Simple File Logger Utility
        /// </summary>
        /// <param name="filePath">File Path Must Be A Log File. If Not Set, Current Executing Assembly Path Will Be Used Instead.</param>
        public FileLoggerUtility(string filePath = null)
        {
            bLoggedFirstMessage = false;

            if (string.IsNullOrEmpty(filePath) == false)
            {
                _fileInfo = new FileInfo(filePath);
            }

            if ((_fileInfo != null && 
                _fileInfo.Directory.Exists && 
                _fileInfo.Extension == ".log") == false)
            {
                string _parentDir = System.IO.Path.GetDirectoryName(new Uri(Assembly.GetExecutingAssembly().CodeBase).LocalPath);
                string _fileName = System.IO.Path.Combine(_parentDir, "FileLoggerUtility.log");
                _fileInfo = new FileInfo(_fileName);
            }    
        }

        /// <summary>
        /// Used To Add String To The Log File.
        /// Similar To Console.WriteLine.
        /// </summary>
        /// <param name="_msg"></param>
        public void AddMessageToLog(string _msg)
        {
            if(_fileInfo != null && 
                _fileInfo.Directory.Exists &&
                string.IsNullOrEmpty(_fileInfo.FullName) == false)
            {
                if (bLoggedFirstMessage)
                {
                    File.AppendAllLines(_fileInfo.FullName, new string[] { _msg });
                }
                else
                {
                    File.WriteAllLines(_fileInfo.FullName, new string[] { _msg });
                }
            }

            bLoggedFirstMessage = true;
        }
    }
}
