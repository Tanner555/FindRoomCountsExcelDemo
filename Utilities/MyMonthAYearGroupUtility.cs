using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyCommonUtilities
{
    /// <summary>
    /// A Simple MonthAYear Grouping Class Which Has A Lot
    /// Of Missing Functionality. Don't Use Without Looking Through Each Method Carefully.
    /// </summary>
    public class MyMonthAYearGroupUtility
    {
        #region Enums
        public enum EDateByMonth
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

        public enum EDateByYear
        {
            Undecided = -1,
            Y2017 = 0, Y2018 = 1, Y2019 = 2, Y2020 = 3, Y2021 = 4, Y2022 = 5, Y2023 = 6, Y2024 = 7, Y2025 = 8
        }
        #endregion

        #region Fields
        public EDateByMonth DateByMonth;
        public EDateByYear DateByYear;
        public int DateByDay;
        #endregion

        #region Initialization
        /// <summary>
        /// Creates A MonthAYear Group Using A Forward-Slash Date Formatted String. Example: 01/01/2021.
        /// Could Also Try Using Standard Date Format. Example: January 1st, 2021.
        /// </summary>
        /// <param name="_date"></param>
        /// <returns></returns>
        public MyMonthAYearGroupUtility(string _date)
        {
            var _dateByMonthAndYear = CalculateDateByMonthAndYear(_date);
            this.DateByMonth = _dateByMonthAndYear.dateByMonth;
            this.DateByYear = _dateByMonthAndYear.dateByYear;
            this.DateByDay = _dateByMonthAndYear.dateByDay;
        }

        private MyMonthAYearGroupUtility()
        {
            this.DateByMonth = EDateByMonth.Undecided;
            this.DateByYear = EDateByYear.Undecided;
        }
        #endregion

        #region Helpers
        public static bool IsMonthMissingDays(MyMonthAYearGroupUtility _monthAYearGroup, int _dayCount)
        {
            return _dayCount != GetDayCountFromMonthAYear(_monthAYearGroup.DateByMonth, _monthAYearGroup.DateByYear);
        }

        public static bool IsMonthMissingDays(EDateByMonth _month, EDateByYear _year, int _dayCount)
        {
            return _dayCount != GetDayCountFromMonthAYear(_month, _year);
        }

        public static int GetDayCountFromMonthAYear(EDateByMonth _month, EDateByYear _year)
        {
            switch (_month)
            {
                case EDateByMonth.Undecided:
                    return -1;
                case EDateByMonth.January:
                    return GetDayCountFromMonthAYear(1, RetrieveYNumberByEYear(_year));
                case EDateByMonth.February:
                    return GetDayCountFromMonthAYear(2, RetrieveYNumberByEYear(_year));
                case EDateByMonth.March:
                    return GetDayCountFromMonthAYear(3, RetrieveYNumberByEYear(_year));
                case EDateByMonth.April:
                    return GetDayCountFromMonthAYear(4, RetrieveYNumberByEYear(_year));
                case EDateByMonth.May:
                    return GetDayCountFromMonthAYear(5, RetrieveYNumberByEYear(_year));
                case EDateByMonth.June:
                    return GetDayCountFromMonthAYear(6, RetrieveYNumberByEYear(_year));
                case EDateByMonth.July:
                    return GetDayCountFromMonthAYear(7, RetrieveYNumberByEYear(_year));
                case EDateByMonth.August:
                    return GetDayCountFromMonthAYear(8, RetrieveYNumberByEYear(_year));
                case EDateByMonth.September:
                    return GetDayCountFromMonthAYear(9, RetrieveYNumberByEYear(_year));
                case EDateByMonth.October:
                    return GetDayCountFromMonthAYear(10, RetrieveYNumberByEYear(_year));
                case EDateByMonth.November:
                    return GetDayCountFromMonthAYear(11, RetrieveYNumberByEYear(_year));
                case EDateByMonth.December:
                    return GetDayCountFromMonthAYear(12, RetrieveYNumberByEYear(_year));
                default:
                    return -1;
            }
        }

        public static int GetDayCountFromMonthAYear(int _month, int _year)
        {
            // Create a new DateTime object with the year and month
            DateTime date = new DateTime(_year, _month, 1);

            // Get the number of days in the month by subtracting the first day of the next month from the first day of this month
            int dayCount = (date.AddMonths(1) - date).Days;

            // Return the day count
            return dayCount;
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
            if (_dateByMonth != EDateByMonth.October &&
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

        static int RetrieveYNumberByEYear(EDateByYear _year)
        {
            if (_year.ToString().Contains("2017"))
                return 2017;
            if (_year.ToString().Contains("2018"))
                return 2018;
            if (_year.ToString().Contains("2019"))
                return 2019;
            if (_year.ToString().Contains("2020"))
                return 2020;
            if (_year.ToString().Contains("2021"))
                return 2021;
            if (_year.ToString().Contains("2022"))
                return 2022;
            if (_year.ToString().Contains("2023"))
                return 2023;
            if (_year.ToString().Contains("2024"))
                return 2024;
            if (_year.ToString().Contains("2025"))
                return 2025;

            return -1;
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
            if (_dateCellValue[1] != '/' &&
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
    }
}
