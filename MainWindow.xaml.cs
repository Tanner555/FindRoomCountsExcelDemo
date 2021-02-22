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

namespace FindRoomCountsExcelDemo
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        #region Enums
        enum EDateByMonth
        {
            Undecided = -1,
            January = 0,
            February = 1,
            March = 2,
            April = 3,
            May = 4,
            June = 5,
            July = 6,
            August = 7,
            September = 8,
            October = 9,
            November = 10,
            December = 11
        }

        enum EDateByYear
        {
            Undecided = -1,
            Y2017 = 0, Y2018 = 1, Y2019 = 2, Y2020 = 3, Y2021 = 4, Y2022 = 5, Y2023 = 6, Y2024 = 7, Y2025 = 8
        }
        #endregion

        #region Fields
        bool bLoggedFirstMessage = false;
        bool bReadingFiles = false;
        bool bIsExecutingRoomCountHandler = false;

        MyFileFinder myReader;
        #endregion

        #region Directories
        public string _desktopDir;
        public string _programFilesDir;
        public string _currentExecutingDirectory;
        public string _currentExeDirLogFile;
        public FileInfo _currentExeDirLogFileInfo;

        void SetupDirectories()
        {
            _desktopDir = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            _programFilesDir = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            _currentExecutingDirectory = System.IO.Path.GetDirectoryName(new Uri(Assembly.GetExecutingAssembly().CodeBase).LocalPath);
            _currentExeDirLogFile = System.IO.Path.Combine(_currentExecutingDirectory, "FileFinderLog.log");
            _currentExeDirLogFileInfo = new FileInfo(_currentExeDirLogFile);
        }
        #endregion

        #region Initialization
        public MainWindow()
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            InitializeComponent();
            SetupDirectories();

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
            if (!string.IsNullOrEmpty(_currentExecutingDirectory) &&
                Directory.Exists(_currentExecutingDirectory) &&
                !string.IsNullOrEmpty(_currentExeDirLogFile) &&
                _currentExeDirLogFileInfo.Directory.Exists)
            {
                if (bLoggedFirstMessage)
                {
                    File.AppendAllLines(_currentExeDirLogFileInfo.FullName, new string[] { _msg });
                }
                else
                {
                    File.WriteAllLines(_currentExeDirLogFileInfo.FullName, new string[] { _msg });
                }
            }

            bLoggedFirstMessage = true;
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

            public EDateByMonth DateByMonth;
            public EDateByYear DateByYear;
            public int DateByDay;

            public DailyRevenueSheetModel(string DateCellValue, string DateCellAddress, 
                int RoomCountCellValue, string RoomCountCellAddress,
                EDateByMonth DateByMonth = EDateByMonth.Undecided,
                EDateByYear DateByYear = EDateByYear.Undecided,
                int DateByDay = -1)
            {
                this.DateCellValue = DateCellValue;
                this.DateCellAddress = DateCellAddress;
                this.RoomCountCellValue = RoomCountCellValue;
                this.RoomCountCellAddress = RoomCountCellAddress;
                this.DateByMonth = DateByMonth;
                this.DateByYear = DateByYear;
                this.DateByDay = DateByDay;
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
                        //Try Finding Room Count Cell End

                        if (string.IsNullOrEmpty(_dateCellAddress) == false)
                        {
                            _dateCellValue = firstSheet.Cells[_dateCellAddress].Text;
                            //SetDebugMessage($"Date Cell Value : {firstSheet.Cells[_dateCellAddress].Text}");
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
                            var _dateByMonthAndYear = CalculateDateByMonthAndYear(_dateCellValue);
                            return new DailyRevenueSheetModel(_dateCellValue, _dateCellAddress, 
                                _myCountTryParse, _roomCountAddress,
                                _dateByMonthAndYear.dateByMonth, 
                                _dateByMonthAndYear.dateByYear,
                                _dateByMonthAndYear.dateByDay);
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

        (EDateByMonth dateByMonth, EDateByYear dateByYear, int dateByDay) CalculateDateByMonthAndYear(string _dateCellValue)
        {
            var _dateByMonth = CalculateDateByMonth(_dateCellValue, out var bIsMonthSpelledOut);
            return (_dateByMonth, CalculateDateByYear(_dateCellValue), 
                CalculateDateByDay(_dateCellValue, bIsMonthSpelledOut, _dateByMonth));
        }

        int CalculateDateByDay(string _dateCellValue, bool bIsMonthSpelledOut, EDateByMonth _dateByMonth)
        {
            int _dateByNum = -1;

            if (bIsMonthSpelledOut)
            {
                return CalculateDayFromDateSpelledOut(_dateCellValue, _dateByMonth.ToString().Length);
            }

            //If Month Is Less Than Two Digits, Use Normal Calculation,
            //Otherwise, Shift the Index Lookup By One.
            if(_dateByMonth != EDateByMonth.October &&
                _dateByMonth != EDateByMonth.November &&
                _dateByMonth != EDateByMonth.December)
            {
                //If 4th Char Has Slash, Day is Single Digit
                if (_dateCellValue[3] == '/' &&
                    int.TryParse(_dateCellValue[2].ToString(), out _dateByNum))
                {
                    return _dateByNum;
                }
                //If 4th Char Isn't A Slash, Day has Two Digits
                if (_dateCellValue[3] != '/' &&
                    int.TryParse(_dateCellValue.Substring(2, 2), out _dateByNum))
                {
                    return _dateByNum;
                }
            }
            else
            {
                //If 5th Char Has Slash, Day is Single Digit
                if (_dateCellValue[4] == '/' &&
                    int.TryParse(_dateCellValue[3].ToString(), out _dateByNum))
                {
                    return _dateByNum;
                }
                //If 5th Char Isn't A Slash, Day has Two Digits
                if (_dateCellValue[4] != '/' &&
                    int.TryParse(_dateCellValue.Substring(3, 2), out _dateByNum))
                {
                    return _dateByNum;
                }
            }

            return _dateByNum;
        }

        int CalculateDayFromDateSpelledOut(string _dateCellValue, int _monthCharCount)
        {
            int _dateByNum = -1;
            int _dayFirstNumIndex = _monthCharCount + 1;
            int _daySecNumIndex = _monthCharCount + 2;
            //If 2nd Date Char Can Be Parsed, Then Date Has Two Digits
            if (int.TryParse(_dateCellValue[_daySecNumIndex].ToString(), out _dateByNum) &&
                int.TryParse(_dateCellValue.Substring(_dayFirstNumIndex, 2), out _dateByNum))
            {
                return _dateByNum;
            }
            //If 1st Char Can Be Parsed And Not The 2nd, Day has One Digit
            if (int.TryParse(_dateCellValue[_dayFirstNumIndex].ToString(), out _dateByNum))
            {
                return _dateByNum;
            }
            return _dateByNum;
        }

        EDateByYear CalculateDateByYear(string _dateCellValue)
        {
            if (_dateCellValue.Contains("2017"))
                return EDateByYear.Y2017;
            if (_dateCellValue.Contains("2018"))
                return EDateByYear.Y2018;
            if (_dateCellValue.Contains("2019"))
                return EDateByYear.Y2019;
            if (_dateCellValue.Contains("2020"))
                return EDateByYear.Y2020;
            if (_dateCellValue.Contains("2021"))
                return EDateByYear.Y2021;
            if (_dateCellValue.Contains("2022"))
                return EDateByYear.Y2022;
            if (_dateCellValue.Contains("2023"))
                return EDateByYear.Y2023;
            if (_dateCellValue.Contains("2024"))
                return EDateByYear.Y2024;
            if (_dateCellValue.Contains("2025"))
                return EDateByYear.Y2025;

            return EDateByYear.Undecided;
        }

        EDateByMonth CalculateDateByMonth(string _dateCellValue, out bool bIsMonthSpelledOut)
        {
            bIsMonthSpelledOut = false;
            if (DateByMonthIsSpelledOut(_dateCellValue, out var _dateByMonth))
            {
                bIsMonthSpelledOut = true;
                return _dateByMonth;
            }            
            int _monthByNum = -1;
            //If 2nd Char Has Slash, Month is Single Digit
            if (_dateCellValue[1] == '/' && 
                int.TryParse(_dateCellValue[0].ToString(), out _monthByNum))
            {
                return RetrieveMonthByNumber(_monthByNum);
            }
            //If 2nd Char Isn't A Slash, Month has Two Digits
            if(_dateCellValue[1] != '/' &&
                int.TryParse(_dateCellValue.Substring(0, 2), out _monthByNum))
            {
                return RetrieveMonthByNumber(_monthByNum);
            }

            return EDateByMonth.Undecided;
        }

        EDateByMonth RetrieveMonthByNumber(int _dateNumber)
        {
            switch (_dateNumber)
            {
                case 1:
                    return EDateByMonth.January;
                case 2:
                    return EDateByMonth.February;
                case 3:
                    return EDateByMonth.March;
                case 4:
                    return EDateByMonth.April;
                case 5:
                    return EDateByMonth.May;
                case 6:
                    return EDateByMonth.June;
                case 7:
                    return EDateByMonth.July;
                case 8:
                    return EDateByMonth.August;
                case 9:
                    return EDateByMonth.September;
                case 10:
                    return EDateByMonth.October;
                case 11:
                    return EDateByMonth.November;
                case 12:
                    return EDateByMonth.December;
                default:
                    return EDateByMonth.Undecided;
            }
        }

        bool DateByMonthIsSpelledOut(string _dateCellValue, out EDateByMonth _dateByMonth)
        {
            _dateByMonth = EDateByMonth.Undecided;
            string _dateLowerCase = _dateCellValue.ToLower();
            if (_dateLowerCase.Contains("january"))
            {
                _dateByMonth = EDateByMonth.January;
                return true;
            }
            if (_dateLowerCase.Contains("february"))
            {
                _dateByMonth = EDateByMonth.February;
                return true;
            }
            if (_dateLowerCase.Contains("march"))
            {
                _dateByMonth = EDateByMonth.March;
                return true;
            }
            if (_dateLowerCase.Contains("april"))
            {
                _dateByMonth = EDateByMonth.April;
                return true;
            }
            if (_dateLowerCase.Contains("may"))
            {
                _dateByMonth = EDateByMonth.May;
                return true;
            }
            if (_dateLowerCase.Contains("june"))
            {
                _dateByMonth = EDateByMonth.June;
                return true;
            }
            if (_dateLowerCase.Contains("july"))
            {
                _dateByMonth = EDateByMonth.July;
                return true;
            }
            if (_dateLowerCase.Contains("august"))
            {
                _dateByMonth = EDateByMonth.August;
                return true;
            }
            if (_dateLowerCase.Contains("september"))
            {
                _dateByMonth = EDateByMonth.September;
                return true;
            }
            if (_dateLowerCase.Contains("october"))
            {
                _dateByMonth = EDateByMonth.October;
                return true;
            }
            if (_dateLowerCase.Contains("november"))
            {
                _dateByMonth = EDateByMonth.November;
                return true;
            }
            if (_dateLowerCase.Contains("december"))
            {
                _dateByMonth = EDateByMonth.December;
                return true;
            }
            return false;
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
                var _myRandom = new System.Random();
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
                                firstSheet.Cells[_myI, 1, _myI, 6].Clear();                                
                                //Iterate At The Beginning of Each Group After Adding Space
                                _myI++;
                            }
                            //Set Random Color And Beginning Iterative Count
                            int _beginningRange = _myI;
                            var _ramColor = GetRandomColor(_myRandom);
                            foreach (var _revenueModel in _revenueSheets)
                            {
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
                                //Iterate After Each Model
                                _myI++;
                            }
                            //Set Ending Iterative Count And Fill Background W/Random Color
                            int _endRange = _myI;
                            firstSheet.Cells[_beginningRange, 1, _endRange, 5].Style.Fill.SetBackground(_ramColor);
                        }
                        //AutoFit Columns And Save To Sheet
                        firstSheet.Column(1).AutoFit();
                        firstSheet.Column(2).AutoFit();
                        firstSheet.Column(3).AutoFit();
                        firstSheet.Column(4).AutoFit();
                        firstSheet.Column(5).AutoFit();
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

        System.Drawing.Color GetRandomColor(System.Random _random)
        {            
            return System.Drawing.Color.FromArgb(_random.Next(0, 255), _random.Next(0, 255), _random.Next(0, 255));
        }
        #endregion

        #region SimpleDurationTimer
        class MySimpleDurationTimer
        {
            TimeSpan _stop;
            TimeSpan _start;

            public MySimpleDurationTimer()
            {
                _start = new TimeSpan(DateTime.Now.Ticks);
            }

            public TimeSpan StopWithDuration()
            {
                _stop = new TimeSpan(DateTime.Now.Ticks);
                return _stop.Subtract(_start);
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

        #region Testing
        /// <summary>
        /// Shouldn't Call Right Now
        /// </summary>
        void TestReadExcel1234()
        {
            using (var package = new ExcelPackage(new FileInfo("Book.xlsx")))
            {
                var firstSheet = package.Workbook.Worksheets["First Sheet"];
                Console.WriteLine("Sheet 1 Data");
                Console.WriteLine($"Cell A2 Value   : {firstSheet.Cells["A2"].Text}");
                Console.WriteLine($"Cell A2 Color   : {firstSheet.Cells["A2"].Style.Font.Color.LookupColor()}");
                Console.WriteLine($"Cell B2 Formula : {firstSheet.Cells["B2"].Formula}");
                Console.WriteLine($"Cell B2 Value   : {firstSheet.Cells["B2"].Text}");
                Console.WriteLine($"Cell B2 Border  : {firstSheet.Cells["B2"].Style.Border.Top.Style}");
                Console.WriteLine("");

                var secondSheet = package.Workbook.Worksheets["Second Sheet"];
                Console.WriteLine($"Sheet 2 Data");
                Console.WriteLine($"Cell A2 Formula : {secondSheet.Cells["A2"].Formula}");
                Console.WriteLine($"Cell A2 Value   : {secondSheet.Cells["A2"].Text}");
            }
        }

        //async Task OLDWriteModelsToOutputSheetNOGrouping(List<DailyRevenueSheetModel> _revenueSheets, FileInfo _outputSheetInfo)
        //{
        //    try
        //    {
        //        var _timer = new MySimpleDurationTimer();
        //        using (var _package = new ExcelPackage(_outputSheetInfo))
        //        {
        //            var firstSheet = _package.Workbook.Worksheets.First();
        //            if (firstSheet != null)
        //            {
        //                int _insertRowIndexRef = 2;
        //                if (firstSheet.Cells.Rows < _revenueSheets.Count + _insertRowIndexRef)
        //                {
        //                    SetDebugMessage("Inserting Rows into output sheet...", false);
        //                    firstSheet.InsertRow(_insertRowIndexRef, _revenueSheets.Count);
        //                }

        //                int _myI = 2;
        //                foreach (var _revenueModel in _revenueSheets)
        //                {
        //                    firstSheet.Cells[_myI, 1].Value = _revenueModel.DateCellValue;
        //                    firstSheet.Cells[_myI, 2].Value = _revenueModel.RoomCountCellValue;
        //                    //firstSheet.Cells[_myI, 3].Value = _revenueModel.DateByMonth.ToString();
        //                    //firstSheet.Cells[_myI, 4].Value = _revenueModel.DateByYear.ToString();
        //                    _myI++;
        //                }
        //                _package.SaveAs(_outputSheetInfo);
        //            }
        //        }

        //        SetDebugMessage("WriteModelsToOutputSheet Duration:" + _timer.StopWithDuration(), false);
        //    }
        //    catch (Exception ex)
        //    {
        //        SetDebugMessage("ERROR: " + ex.Message);
        //    }
        //}
        #endregion
    }
}
