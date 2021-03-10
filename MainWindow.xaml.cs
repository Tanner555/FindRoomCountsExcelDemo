using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.ComponentModel;
using System.IO;
using System.Reflection;
using OfficeOpenXml;
using Microsoft.Win32;
using Ookii.Dialogs.Wpf;
using MyCommonUtilities;
using EDateByMonth = MyCommonUtilities.MyMonthAYearGroupUtility.EDateByMonth;
using EDateByYear = MyCommonUtilities.MyMonthAYearGroupUtility.EDateByYear;

namespace FindRoomCountsExcelDemo
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        #region Properties
        FileLoggerUtility myFileLoggerUtility
        {
            get
            {
                if(_myFileLoggerUtility == null)
                {
                    _myFileLoggerUtility = new FileLoggerUtility();
                }
                return _myFileLoggerUtility;
            }
        }
        FileLoggerUtility _myFileLoggerUtility = null;
        #endregion

        #region Fields
        bool bReadingFiles = false;
        bool bIsExecutingRoomCountHandler = false;

        MyFileFinder myReader;
        #endregion

        #region Initialization
        public MainWindow()
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            InitializeComponent();

            bIsExecutingRoomCountHandler = false;
            bReadingFiles = false;

            btn_findAndOutputRoomCounts.Click += btn_FindAndOutputRoomCounts_Click;
            btnFindDailyRevFolder.Click += btn_FindDailyRevFolder_Click;
            btnFindOutputRevenueFile.Click += btn_FindOutputRevenueFile_Click;

            SetDailyRevFolderPathToSettingIfValid();
            SetOutputRevFilePathToSettingIfValid();
        }
        #endregion

        #region Handlers
        private async void btn_FindAndOutputRoomCounts_Click(object sender, RoutedEventArgs e)
        {
            //Prevents Executing More Than Once if Clicked Multiple Times
            if (bIsExecutingRoomCountHandler)
            {
                SetDebugMessage("Currently Calculating or Outputing Room Counts, Cannot Perform Action.");
            }

            bIsExecutingRoomCountHandler = true;

            string _errnoMsg = "";
            DirectoryInfo _dailyRevenueFolderInfo = new DirectoryInfo(DailyRevenueFolderPath.Text);
            FileInfo _outputRevenueFileInfo = new FileInfo(OutputRevenueFilePath.Text);

            if (_dailyRevenueFolderInfo.Exists == false)
            {
                _errnoMsg = "Daily Revenue Folder Path is Either Not Set or Not Valid, Cannot Continue.";
                MessageBox.Show(_errnoMsg);
                SetDebugMessage(_errnoMsg);
                return;
            }

            if (_outputRevenueFileInfo.Exists == false ||
                _outputRevenueFileInfo.Length == 0 ||
                _outputRevenueFileInfo.Extension.Contains("xlsx") == false)
            {
                _errnoMsg = "Output Revenue File Path is Either Not Set or Not Valid, Cannot Continue.";
                MessageBox.Show(_errnoMsg);
                SetDebugMessage(_errnoMsg);
                return;
            }

            SetDebugMessage($"Finding Files From Revenue Folder.....");
            var _files = await ReadAllFiles(_dailyRevenueFolderInfo.FullName);
            SetDebugMessage($"Reading Data From {myReader.filesFound} Excel Files");
            var _revenueModels = await ReadAndCalculateDataFromExcelFiles(_files);
            SetDebugMessage($"Grouping Rev Models Month/Year.....");
            var _modelsGroupedIntoMonthAYear = await PutRevModelsIntoMonthAYearGroups(_revenueModels);
            SetDebugMessage($"Writing {_revenueModels.Count} Rev Models To Output Sheet");
            await WriteModelsToOutputSheet(_modelsGroupedIntoMonthAYear, _outputRevenueFileInfo, _revenueModels.Count);
            //Once Executing This Func is All Done
            SetDebugMessage("All done writing revenue data to output excel sheet");
            bIsExecutingRoomCountHandler = false;
        }

        private void btn_FindDailyRevFolder_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var _dialog = new VistaFolderBrowserDialog();
                var _dialogSuccessful = _dialog.ShowDialog(this);
                if (_dialogSuccessful != null && _dialogSuccessful.Value)
                {
                    DailyRevenueFolderPath.Text = _dialog.SelectedPath;
                    SetDailyRevFolderPathSetting(_dialog.SelectedPath);
                    SetDebugMessage($"DailyRevFolderPath Successfully Changed.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                SetDebugMessage(ex.Message);
            }
        }

        private void btn_FindOutputRevenueFile_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                VistaOpenFileDialog ofd = new VistaOpenFileDialog() { Filter = "Excel Sheet (*.xlsx)|*.xlsx", ValidateNames = true, Multiselect = false };
                if (ofd.ShowDialog().Value == true)
                {
                    OutputRevenueFilePath.Text = ofd.FileName;
                    SetOutputRevFilePathSetting(ofd.FileName);
                    SetDebugMessage($"OutputRevenueFilePath Successfully Changed.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                SetDebugMessage(ex.Message);
            }
        }

        void UpdateFileStatTextsHandler()
        {
            UpdateFileStatTextsFromReader();
        }

        void OnShowErrorMessageFromReaderHandler(string _msg)
        {
            SetDebugMessage(_msg);
        }
        #endregion

        #region ModelSetters
        void UpdateFileStatTextsFromReader()
        {
            if (myReader == null) return;
            SetFilesRead(myReader.filesRead);
            SetFilesFound(myReader.filesFound);
            SetDirsFound(myReader.directoriesFound);
        }

        void SetRevenueFilesRead(int _value)
        {
            RevenueFilesReadStat.Text = $"{_value}";
        }

        void SetFilesRead(int _value)
        {
            FilesReadStat.Text = $"{_value}";
        }

        void SetFilesFound(int _value)
        {
            FilesFoundStat.Text = $"{_value}";
        }

        void SetDirsFound(int _value)
        {
            DirsFoundStat.Text = $"{_value}";
        }

        void SetDebugMessage(string _msg, bool bUpdateText = true)
        {
            if (bUpdateText)
            {
                DebugMessages.Text = $"{_msg}";
            }

            myFileLoggerUtility.AddMessageToLog(_msg);
        }
        #endregion

        #region ReadFiles
        async Task<List<string>> ReadAllFiles(string _dir, Func<string, bool> _filePathCondition = null)
        {
            var _timer = new MySimpleDurationTimer();
            if (bReadingFiles)
            {
                SetDebugMessage("Cannot Read Files Until Current Folder is Read");
                return null;
            }
            bReadingFiles = true;

            myReader = new MyFileFinder();
            myReader.OnShowErrorMessage += OnShowErrorMessageFromReaderHandler;

            if (_filePathCondition == null)
            {
                _filePathCondition = new Func<string, bool>(
                    (_fPath) =>
                    {
                        FileInfo _info = new FileInfo(_fPath);
                        if (!_info.Exists)
                        {
                            SetDebugMessage($"File {_fPath} cannot be found - file doesn't exist");
                            return false;
                        }

                        return _info.Extension == ".xlsx";
                    });
            }

            InvokeTimer.CancelInvoke();
            InvokeTimer.InvokeRepeating(100f, UpdateFileStatTextsHandler, TaskScheduler.FromCurrentSynchronizationContext());

            string[] _filesArray = await myReader.GetReadFromDirTask(_dir, _filePathCondition);

            bReadingFiles = false;
            myReader.OnShowErrorMessage -= OnShowErrorMessageFromReaderHandler;

            UpdateFileStatTextsFromReader();

            InvokeTimer.CancelInvoke();
            SetDebugMessage("ReadAllFiles Duration:" + _timer.StopWithDuration(), false);
            return _filesArray != null ? _filesArray.ToList() : null;
        }
        #endregion

        #region ExcelSheetModel
        class DailyRevenueSheetModel
        {
            [OfficeOpenXml.Attributes.EpplusTableColumn(Header = "Date")]
            public string DateCellValue;
            [OfficeOpenXml.Attributes.EpplusIgnore]
            public string DateCellAddress;
            [OfficeOpenXml.Attributes.EpplusTableColumn(Header = "Room Count")]
            public int RoomCountCellValue;
            [OfficeOpenXml.Attributes.EpplusIgnore]
            public string RoomCountCellAddress;

            private MyMonthAYearGroupUtility _myMonthAYearGroupUtility;

            public string SheetFileName;
            public string SheetParentFolder;

            public int DateByDay => _myMonthAYearGroupUtility.DateByDay;
            public EDateByMonth DateByMonth => _myMonthAYearGroupUtility.DateByMonth;
            public EDateByYear DateByYear => _myMonthAYearGroupUtility.DateByYear;

            public DailyRevenueSheetModel(string DateCellValue, string DateCellAddress, 
                int RoomCountCellValue, string RoomCountCellAddress,
                string dateWForwardSlashesStrictFormatted,
                string SheetFileName = "", string SheetParentFolder = "")
            {
                this.DateCellValue = DateCellValue;
                this.DateCellAddress = DateCellAddress;
                this.RoomCountCellValue = RoomCountCellValue;
                this.RoomCountCellAddress = RoomCountCellAddress;
                this._myMonthAYearGroupUtility = new MyMonthAYearGroupUtility(dateWForwardSlashesStrictFormatted);
                this.SheetFileName = SheetFileName;
                this.SheetParentFolder = SheetParentFolder;
            }
        }
        #endregion

        #region ReadingExcel
        async Task<List<DailyRevenueSheetModel>> ReadAndCalculateDataFromExcelFiles(List<string> _excelFiles)
        {
            var _timer = new MySimpleDurationTimer();
            if (_excelFiles == null || _excelFiles.Count <= 0)
            {
                SetDebugMessage("No Excel Files To Calculate...");
                return null;
            }

            List<DailyRevenueSheetModel> _revenueSheets = new List<DailyRevenueSheetModel>();

            foreach (var _excelFile in _excelFiles)
            {
                var _revenueSheet = await FindDataFromExcelFile(_excelFile);
                if (_revenueSheet != null)
                {
                    _revenueSheets.Add(_revenueSheet);
                }
                else
                {
                    SetDebugMessage($"Couldn't Add Revenue Sheet {_excelFile}.", false);
                }
            }

            SetRevenueFilesRead(_revenueSheets.Count);
            SetDebugMessage("ReadAndCalculateDataFromExcelFiles Duration:" + _timer.StopWithDuration(), false);
            return _revenueSheets;
        }

        async Task<DailyRevenueSheetModel> FindDataFromExcelFile(string _excelFile)
        {
            FileInfo _excelFileInfo = new FileInfo(_excelFile);
            if (_excelFile.Length > 0 && _excelFileInfo.Exists && _excelFileInfo.Extension.Contains("xlsx"))
            {
                using (var _package = new ExcelPackage(_excelFileInfo))
                {
                    var firstSheet = _package.Workbook.Worksheets[0];
                    if (firstSheet != null)
                    {
                        string _dateCellValue = "";
                        string _dateCellAddress = "";
                        string _roomCountValue = "";
                        string _roomCountAddress = "";
                        if (firstSheet.Cells["I1"].Text.ToLower().Contains("date"))
                        {
                            _dateCellAddress = "J1";
                        }
                        else
                        {
                            //Try To Find Cell With Date:
                            foreach (var _cell in firstSheet.Cells)
                            {
                                if (_cell.Text.ToLower().Contains("date"))
                                {
                                    _dateCellAddress = firstSheet.Cells[_cell.Rows, _cell.Columns + 1].Address;
                                }
                            }
                        }

                        //Try Finding Room Count Cell Start
                        if (firstSheet.Cells["P1"].Text.ToLower().Contains("count") &&
                            firstSheet.Cells["P1"].Merge && firstSheet.Cells["Q1"].Merge &&
                            firstSheet.Cells["R1"].Merge == false)
                        {
                            _roomCountAddress = "R1";
                        }
                        else if (firstSheet.Cells["Q1"].Text.ToLower().Contains("count") &&
                            firstSheet.Cells["Q1"].Merge && firstSheet.Cells["R1"].Merge &&
                            firstSheet.Cells["S1"].Merge == false)
                        {
                            _roomCountAddress = "S1";
                        }
                        else if (firstSheet.Cells["O1"].Text.ToLower().Contains("count") &&
                            firstSheet.Cells["O1"].Merge && firstSheet.Cells["P1"].Merge &&
                            firstSheet.Cells["Q1"].Merge == false)
                        {
                            _roomCountAddress = "Q1";
                        }
                        else if (firstSheet.Cells["R1"].Text.ToLower().Contains("count") &&
                            firstSheet.Cells["R1"].Merge && firstSheet.Cells["S1"].Merge &&
                            firstSheet.Cells["T1"].Merge == false)
                        {
                            _roomCountAddress = "T1";
                        }

                        //Validating Addresses/Values And Returning New Model
                        if (string.IsNullOrEmpty(_dateCellAddress) == false)
                        {
                            _dateCellValue = firstSheet.Cells[_dateCellAddress].Text;                            
                        }

                        if (string.IsNullOrEmpty(_roomCountAddress) == false)
                        {
                            _roomCountValue = firstSheet.Cells[_roomCountAddress].Text;
                        }
                        int _myCountTryParse = -1;
                        if (string.IsNullOrEmpty(_dateCellValue) == false &&
                            string.IsNullOrEmpty(_dateCellAddress) == false &&
                            string.IsNullOrEmpty(_roomCountValue) == false &&
                            string.IsNullOrEmpty(_roomCountAddress) == false &&
                            int.TryParse(_roomCountValue, out _myCountTryParse) && _myCountTryParse != -1)
                        {
                            return new DailyRevenueSheetModel(_dateCellValue, _dateCellAddress, 
                                _myCountTryParse, _roomCountAddress, _dateCellValue,
                                _excelFileInfo.Name, _excelFileInfo.Directory.Name);
                        }
                    }
                }
            }
            else
            {
                SetDebugMessage($"Couldn't Find Spreadsheet at: {_excelFileInfo.FullName}");
            }
            return null;
        }        
        #endregion

        #region PutRevModelsIntoMonthAYearGroups
        async Task<Dictionary<string, List<DailyRevenueSheetModel>>> PutRevModelsIntoMonthAYearGroups(List<DailyRevenueSheetModel> _revenueSheets)
        {
            var _timer = new MySimpleDurationTimer();
            var _yearAMonthRevGroups = new Dictionary<string, List<DailyRevenueSheetModel>>();

            var _organizedRevenueSheets = from _sheet in _revenueSheets
                                         orderby _sheet.DateByDay ascending
                                         orderby _sheet.DateByMonth descending
                                         orderby _sheet.DateByYear descending
                                         select _sheet;            

            //Iterate Over Every Sheet, Creating New MonthAYear Groups As Needed
            foreach (var _revenueSheet in _organizedRevenueSheets)
            {
                string _revenueKey = $"{_revenueSheet.DateByMonth}-{_revenueSheet.DateByYear}";
                if (_yearAMonthRevGroups.ContainsKey(_revenueKey))
                {
                    //Contains Key, Add Rev Model To The List
                    var _revSheetGroupByKey = _yearAMonthRevGroups[_revenueKey];
                    if (_revSheetGroupByKey != null)
                    {
                        _revSheetGroupByKey.Add(_revenueSheet);
                    }
                }
                else
                {
                    //Doesn't Contain Key, Create New List W/Model And Add To Dictionary
                    var _newRevenueGroupFromKey = new List<DailyRevenueSheetModel>();
                    _newRevenueGroupFromKey.Add(_revenueSheet);
                    _yearAMonthRevGroups.Add(_revenueKey, _newRevenueGroupFromKey);
                }
            }
            SetDebugMessage($"Organized Sheet Count: {_organizedRevenueSheets.Count()}", false);
            SetDebugMessage("PutRevModelsIntoMonthAYearGroups Duration:" + _timer.StopWithDuration(), false);
            SetDebugMessage($"Rev MonthAYear Groups Created: {_yearAMonthRevGroups.Count}", false);            
            return _yearAMonthRevGroups;
        }
        #endregion

        #region WritingToOutputSheet
        async Task WriteModelsToOutputSheet(Dictionary<string, List<DailyRevenueSheetModel>> _modelsGroupedIntoMonthAYear, FileInfo _outputSheetInfo, int _revModelCount)
        {
            try
            {
                var _timer = new MySimpleDurationTimer();
                var _colorUtil = new SimpleColorUtility();
                int _revModelPlusGroupsCount =
                    _revModelCount + _modelsGroupedIntoMonthAYear.Count + 2;
                using (var _package = new ExcelPackage(_outputSheetInfo))
                {
                    var firstSheet = _package.Workbook.Worksheets.First();
                    if (firstSheet != null)
                    {
                        //Clear And Insert Rows As Needed
                        firstSheet.Cells.Clear();
                        int _insertRowIndexRef = 2;
                        if (firstSheet.Cells.Rows < _revModelPlusGroupsCount + _insertRowIndexRef)
                        {
                            SetDebugMessage("Inserting Rows into output sheet...", false);
                            firstSheet.InsertRow(_insertRowIndexRef, _revModelPlusGroupsCount);
                        }

                        //Add Headers
                        firstSheet.Cells[1, 1].Value = "Date";
                        firstSheet.Cells[1, 1].Style.Font.Italic = true;
                        firstSheet.Cells[1, 1].Style.Font.Size = 14.0f;
                        firstSheet.Cells[1, 1].Style.Font.UnderLine = true;
                        firstSheet.Cells[1, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;                        
                        firstSheet.Cells[1, 2].Value = "Room Count";
                        firstSheet.Cells[1, 2].Style.Font.Italic = true;
                        firstSheet.Cells[1, 2].Style.Font.Size = 14.0f;
                        firstSheet.Cells[1, 2].Style.Font.UnderLine = true;
                        firstSheet.Cells[1, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        //Iterate Through Groups And Write To Sheet
                        int _myI = 2;
                        foreach (var (_revGroupKey, _revenueSheets) in _modelsGroupedIntoMonthAYear)
                        {
                            //Only Add Space After The First Group Finishes
                            if (_myI > 2)
                            {
                                //Add Empty Space After Every Group
                                firstSheet.Cells[_myI, 1, _myI, 8].Clear();                                
                                //Iterate At The Beginning of Each Group After Adding Space
                                _myI++;
                            }
                            //Set Random Color And Beginning Iterative Count
                            int _beginningRange = _myI;
                            var _ramColor = _colorUtil.GetRandomColor();
                            //Monthly Room Count And DateByMonthAYear
                            int _monthlyRoomCount = 0;                            
                            EDateByMonth _myDateByMonth = EDateByMonth.Undecided;
                            EDateByYear _myDateByYear = EDateByYear.Undecided;
                            //Iterate Through Revenue Sheets
                            foreach (var _revenueModel in _revenueSheets)
                            {
                                //Figure Out Which MonthAYear We're On If Undecided
                                if(_myDateByMonth == EDateByMonth.Undecided)
                                {
                                    _myDateByMonth = _revenueModel.DateByMonth;
                                }
                                if(_myDateByYear == EDateByYear.Undecided)
                                {
                                    _myDateByYear = _revenueModel.DateByYear;
                                }
                                //Add All Room Counts To Monthly Room Count
                                _monthlyRoomCount += _revenueModel.RoomCountCellValue;
                                //Date Cell
                                firstSheet.Cells[_myI, 1].Value = _revenueModel.DateCellValue;
                                firstSheet.Cells[_myI, 1].Style.Font.UnderLine = true;
                                firstSheet.Cells[_myI, 1].Style.Font.Bold = true;
                                firstSheet.Cells[_myI, 1].Style.Font.Size = 16.0f;
                                firstSheet.Cells[_myI, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                //Room Count Cell
                                firstSheet.Cells[_myI, 2].Value = _revenueModel.RoomCountCellValue;
                                firstSheet.Cells[_myI, 2].Style.Font.UnderLine = true;
                                firstSheet.Cells[_myI, 2].Style.Font.Bold = true;
                                firstSheet.Cells[_myI, 2].Style.Font.Size = 18.0f;
                                firstSheet.Cells[_myI, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                //Debugging Month, Year and Date Grouping
                                firstSheet.Cells[_myI, 3].Value = _revenueModel.DateByMonth.ToString();
                                firstSheet.Cells[_myI, 3].Style.Font.Size = 11.0f;
                                firstSheet.Cells[_myI, 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                firstSheet.Cells[_myI, 4].Value = _revenueModel.DateByYear.ToString();
                                firstSheet.Cells[_myI, 4].Style.Font.Size = 11.0f;
                                firstSheet.Cells[_myI, 4].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                firstSheet.Cells[_myI, 5].Value = _revenueModel.DateByDay;
                                firstSheet.Cells[_myI, 5].Style.Font.Size = 11.0f;
                                firstSheet.Cells[_myI, 5].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                //Also Debugging FileName And Parent Directory
                                firstSheet.Cells[_myI, 6].Value = _revenueModel.SheetFileName;
                                firstSheet.Cells[_myI, 6].Style.Font.Size = 11.0f;
                                firstSheet.Cells[_myI, 6].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                firstSheet.Cells[_myI, 7].Value = _revenueModel.SheetParentFolder;
                                firstSheet.Cells[_myI, 7].Style.Font.Size = 11.0f;
                                firstSheet.Cells[_myI, 7].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                //Iterate After Each Model
                                _myI++;                                
                            }
                            //Only Show Monthly Room Count if it's Greater than 0
                            //And There's More Than 5 Revenue Sheets In The Monthly Group
                            if (_monthlyRoomCount > 0 && _revenueSheets.Count > 5)
                            {
                                //If Month Is Missing Days, Then Add Missing Notifier.
                                if (MyMonthAYearGroupUtility.IsMonthMissingDays(_myDateByMonth, _myDateByYear, _revenueSheets.Count))
                                {
                                    firstSheet.Cells[_myI - 3, 8].Value = "Month Missing Days.";
                                    firstSheet.Cells[_myI - 3, 8].Style.Font.UnderLine = true;
                                    firstSheet.Cells[_myI - 3, 8].Style.Font.Italic = true;
                                    firstSheet.Cells[_myI - 3, 8].Style.Font.Size = 14.0f;
                                    firstSheet.Cells[_myI - 3, 8].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                }
                                //Monthly Count Header On Row-Day Before Last, And Before Iteration
                                firstSheet.Cells[_myI - 2, 8].Value = "Monthly Count";
                                firstSheet.Cells[_myI - 2, 8].Style.Font.UnderLine = true;
                                firstSheet.Cells[_myI - 2, 8].Style.Font.Italic = true;
                                firstSheet.Cells[_myI - 2, 8].Style.Font.Size = 14.0f;
                                firstSheet.Cells[_myI - 2, 8].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                //Show Monthly Room Count on Last Row-Day Before Iteration
                                firstSheet.Cells[_myI - 1, 8].Value = _monthlyRoomCount;
                                firstSheet.Cells[_myI - 1, 8].Style.Font.UnderLine = true;
                                firstSheet.Cells[_myI - 1, 8].Style.Font.Bold = true;
                                firstSheet.Cells[_myI - 1, 8].Style.Font.Size = 18.0f;
                                firstSheet.Cells[_myI - 1, 8].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            }
                            //Set Ending Iterative Count And Fill Background W/Random Color
                            int _endRange = _myI;
                            firstSheet.Cells[_beginningRange, 1, _endRange, 8].Style.Fill.SetBackground(_ramColor);
                        }
                        //AutoFit Columns And Save To Sheet
                        firstSheet.Column(1).AutoFit();
                        firstSheet.Column(2).AutoFit();
                        firstSheet.Column(3).AutoFit();
                        firstSheet.Column(4).AutoFit();
                        firstSheet.Column(5).AutoFit();
                        firstSheet.Column(6).AutoFit();
                        firstSheet.Column(7).AutoFit();
                        firstSheet.Column(8).AutoFit();
                        _package.SaveAs(_outputSheetInfo);
                    }
                }

                SetDebugMessage("WriteModelsToOutputSheet Duration:" + _timer.StopWithDuration(), false);
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR: " + ex.Message);
                SetDebugMessage("ERROR: " + ex.Message);
            }
        }
        #endregion

        #region SettingToMainWindowHelpers
        void SetDailyRevFolderPathToSettingIfValid()
        {
            string _dailyRevSetting = GetDailyRevFolderPathSetting();
            if (string.IsNullOrEmpty(_dailyRevSetting) == false)
            {
                DailyRevenueFolderPath.Text = _dailyRevSetting;
            }
        }

        void SetOutputRevFilePathToSettingIfValid()
        {
            string _outFileSetting = GetOutputRevFilePathSetting();
            if (string.IsNullOrEmpty(_outFileSetting) == false)
            {
                OutputRevenueFilePath.Text = _outFileSetting;
            }
        }
        #endregion

        #region SettingHelpers
        string GetDailyRevFolderPathSetting()
        {
            return Properties.Settings.Default.DailyRevFolderPathSetting;
        }

        void SetDailyRevFolderPathSetting(string _folderPath)
        {
            Properties.Settings.Default.DailyRevFolderPathSetting = _folderPath;
            Properties.Settings.Default.Save();
        }

        string GetOutputRevFilePathSetting()
        {
            return Properties.Settings.Default.OutputRevFilePathSetting;
        }

        void SetOutputRevFilePathSetting(string _filePath)
        {
            Properties.Settings.Default.OutputRevFilePathSetting = _filePath;
            Properties.Settings.Default.Save();
        }
        #endregion

    }
}
