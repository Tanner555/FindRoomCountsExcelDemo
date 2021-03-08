using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace MyCommonUtilities
{
    public class MyFileFinder
    {
        #region Events
        public delegate void GenEventHandler(int _stat);
        public event GenEventHandler OnFilesReadUpdate;
        public event GenEventHandler OnFilesFoundUpdate;
        public event GenEventHandler OnDirectoriesFoundUpdate;

        public delegate void OneStringHandler(string _msg);
        public event OneStringHandler OnShowErrorMessage;

        private void CallOnShowErrorMessage(string _msg)
        {
            if (OnShowErrorMessage != null) OnShowErrorMessage(_msg);
        }
        #endregion

        #region Properties
        public int filesRead
        {
            get { return _filesRead; }
            protected set
            {
                _filesRead = value;
                if (OnFilesReadUpdate != null) OnFilesReadUpdate(_filesRead);
            }
        }
        public int filesFound
        {
            get { return _filesFound; }
            protected set
            {
                _filesFound = value;
                if (OnFilesFoundUpdate != null) OnFilesFoundUpdate(_filesFound);
            }
        }
        public int directoriesFound
        {
            get { return _directoriesFound; }
            protected set
            {
                _directoriesFound = value;
                if (OnDirectoriesFoundUpdate != null) OnDirectoriesFoundUpdate(_directoriesFound);
            }
        }
        #endregion

        #region Fields
        int _filesRead = 0;
        int _filesFound = 0;
        int _directoriesFound = 0;

        string[] _fileList = new string[0];
        #endregion

        #region Initialization
        public MyFileFinder()
        {
            InitializeFinder();
        }

        void InitializeFinder()
        {
            filesRead = 0;
            filesFound = 0;
            directoriesFound = 0;
            _fileList = new string[0];
        }
        #endregion

        //Can be called by any class
        public async Task<string[]> ReadFromDirectory(string _dir, Func<string, bool> _filePathCondition = null)
        {
            InitializeFinder();
            if (Directory.Exists(_dir))
            {
                return await GetAllFilesAsync(_dir, _filePathCondition);
            }
            else
            { 
                return null;
                //throw new SystemException($"Cannot find file from: ${_dir}");
            }
        }

        //Can be called by any class
        public Task<string[]> GetReadFromDirTask(string _dir, Func<string, bool> _filePathCondition = null)
        {
            InitializeFinder();
            if (Directory.Exists(_dir))
            {
                //return await GetAllFilesAsync(_dir, _filePathCondition);
                return GetAllFilesWrapper(_dir, _filePathCondition);
            }
            else
            {
                return null;
                //throw new SystemException($"Cannot find file from: ${_dir}");
            }
        }

        private Task<string[]> GetAllFilesWrapper(string _dir, Func<string, bool> _filePathCondition = null)
        {
            return Task.Factory.StartNew(() =>
            GetAllFiles(_dir, _filePathCondition), CancellationToken.None, TaskCreationOptions.AttachedToParent,TaskScheduler.Current);
        }

        private string[] GetAllFiles(string _dir, Func<string, bool> _filePathCondition = null)
        {
            List<String> files = new List<String>();
            try
            {
                foreach (string f in Directory.GetFiles(_dir))
                {
                    filesRead++;
                    if (_filePathCondition != null && _filePathCondition(f))
                    {
                        filesFound++;
                        files.Add(f);
                    }
                }
                foreach (string d in Directory.GetDirectories(_dir))
                {
                    directoriesFound++;
                    //Thread.Sleep(500);
                    files.AddRange(GetAllFiles(d, _filePathCondition));
                }
            }
            catch (System.Exception excpt)
            {
                CallOnShowErrorMessage(excpt.Message);
            }

            return files.ToArray();
        }

        private async Task<string[]> GetAllFilesAsync(string _dir, Func<string, bool> _filePathCondition = null)
        {
            List<String> files = new List<String>();
            try
            {
                foreach (string f in Directory.GetFiles(_dir))
                {
                    filesRead++;
                    if (_filePathCondition != null && _filePathCondition(f))
                    {
                        filesFound++;
                        files.Add(f);
                    }                   
                }
                foreach (string d in Directory.GetDirectories(_dir))
                {
                    directoriesFound++;
                    //Thread.Sleep(500);
                    files.AddRange(await GetAllFilesAsync(d, _filePathCondition));
                }
            }
            catch (System.Exception excpt)
            {
                CallOnShowErrorMessage(excpt.Message);
            }

            return files.ToArray();
        }
    }
}
