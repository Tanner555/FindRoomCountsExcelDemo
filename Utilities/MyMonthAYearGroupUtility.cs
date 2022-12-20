using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyCommonUtilities
{
    /// <summary>
    /// A Simple MonthAYear Grouping Class.
    /// </summary>
    public class MyMonthAYearGroupUtility
    {
        #region Fields
        public int DateByMonth;
        public int DateByYear;
        public int DateByDay;
        public string MonthSpelledOut;
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
            (var _day, var _month, var _year) = CalculateDayMonthAYearByDateFormatted(_date);
            this.DateByMonth = _month;
            this.DateByYear = _year;
            this.DateByDay = _day;
            MonthSpelledOut = GetMonthSpelledOut(_month);
        }

        private MyMonthAYearGroupUtility()
        {
            this.DateByMonth = -1;
            this.DateByYear = -1;
            this.DateByDay = -1;
            this.MonthSpelledOut = "";
        }
        #endregion

        #region Helpers
        /// <summary>
        /// Figures Out if Month is Missing Days Based on the Current Day Count / Number.
        /// </summary>
        public static bool IsMonthMissingDays(int _month, int _year, int _dayCount)
        {
            return _dayCount != GetDayCountFromMonthAYear(_month, _year);
        }

        public static int GetDayCountFromMonthAYear(int _month, int _year)
        {
            //If Month/Year is Undecided, then Do Not Try To Create a New DateTime
            if (_month == -1 || _year == -1) return -1;

            // Create a new DateTime object with the year and month
            DateTime date = new DateTime(_year, _month, 1);

            // Get the number of days in the month by subtracting the first day of the next month from the first day of this month
            int dayCount = (date.AddMonths(1) - date).Days;

            // Return the day count
            return dayCount;
        }

        /// <summary>
        /// Calculate The Year By A Date Formatted (mm/dd/yy)
        /// </summary>
        /// <param name="dateFormatted">Date Formatted By Month/Day/Year (mm/dd/yy), IE: 12/19/2022</param>
        /// <returns>Day, Month and Year In Number Form</returns>
        (int _day, int _month, int _year) CalculateDayMonthAYearByDateFormatted(string dateFormatted)
        {
            // Define a custom format provider that includes a leading zero for single-digit days
            var formatProvider = new MyLeadingZeroDateFormatProvider();

            // Try to parse the date string using the custom format provider
            if (DateTime.TryParse(dateFormatted, formatProvider, DateTimeStyles.None, out DateTime date))
            {
                // If the parsing was successful, return the month and year
                return (date.Day, date.Month, date.Year);
            }
            else
            {
                // If the parsing was not successful, return -1
                return (-1, -1, -1);
            }
        }

        string GetMonthSpelledOut(int _month)
        {
            switch (_month)
            {
                case 1:
                    return "January";
                case 2:
                    return "February";
                case 3:
                    return "March";
                case 4:
                    return "April";
                case 5:
                    return "May";
                case 6:
                    return "June";
                case 7:
                    return "July";
                case 8:
                    return "August";
                case 9:
                    return "September";
                case 10:
                    return "October";
                case 11:
                    return "November";
                case 12:
                    return "December";
                default:
                    return "";
            }
        }
        #endregion

        #region LegacyCode
        //public enum EDateByMonth
        //{
        //    Undecided = -1,
        //    January = 0,
        //    February = 1,
        //    March = 2,
        //    April = 3,
        //    May = 4,
        //    June = 5,
        //    July = 6,
        //    August = 7,
        //    September = 8,
        //    October = 9,
        //    November = 10,
        //    December = 11
        //}

        //bool DateByMonthIsSpelledOut(string _dateCellValue, out EDateByMonth _dateByMonth)
        //{
        //    _dateByMonth = EDateByMonth.Undecided;
        //    string _dateLowerCase = _dateCellValue.ToLower();
        //    if (_dateLowerCase.Contains("january"))
        //    {
        //        _dateByMonth = EDateByMonth.January;
        //        return true;
        //    }
        //    if (_dateLowerCase.Contains("february"))
        //    {
        //        _dateByMonth = EDateByMonth.February;
        //        return true;
        //    }
        //    if (_dateLowerCase.Contains("march"))
        //    {
        //        _dateByMonth = EDateByMonth.March;
        //        return true;
        //    }
        //    if (_dateLowerCase.Contains("april"))
        //    {
        //        _dateByMonth = EDateByMonth.April;
        //        return true;
        //    }
        //    if (_dateLowerCase.Contains("may"))
        //    {
        //        _dateByMonth = EDateByMonth.May;
        //        return true;
        //    }
        //    if (_dateLowerCase.Contains("june"))
        //    {
        //        _dateByMonth = EDateByMonth.June;
        //        return true;
        //    }
        //    if (_dateLowerCase.Contains("july"))
        //    {
        //        _dateByMonth = EDateByMonth.July;
        //        return true;
        //    }
        //    if (_dateLowerCase.Contains("august"))
        //    {
        //        _dateByMonth = EDateByMonth.August;
        //        return true;
        //    }
        //    if (_dateLowerCase.Contains("september"))
        //    {
        //        _dateByMonth = EDateByMonth.September;
        //        return true;
        //    }
        //    if (_dateLowerCase.Contains("october"))
        //    {
        //        _dateByMonth = EDateByMonth.October;
        //        return true;
        //    }
        //    if (_dateLowerCase.Contains("november"))
        //    {
        //        _dateByMonth = EDateByMonth.November;
        //        return true;
        //    }
        //    if (_dateLowerCase.Contains("december"))
        //    {
        //        _dateByMonth = EDateByMonth.December;
        //        return true;
        //    }
        //    return false;
        //}

        //int CalculateDateByDay(string _dateCellValue, bool bIsMonthSpelledOut, EDateByMonth _dateByMonth)
        //{
        //    int _dateByNum = -1;

        //    if (bIsMonthSpelledOut)
        //    {
        //        return CalculateDayFromDateSpelledOut(_dateCellValue, _dateByMonth.ToString().Length);
        //    }

        //    //If Month Is Less Than Two Digits, Use Normal Calculation,
        //    //Otherwise, Shift the Index Lookup By One.
        //    if (_dateByMonth != EDateByMonth.October &&
        //        _dateByMonth != EDateByMonth.November &&
        //        _dateByMonth != EDateByMonth.December)
        //    {
        //        //If 4th Char Has Slash, Day is Single Digit
        //        if (_dateCellValue[3] == '/' &&
        //            int.TryParse(_dateCellValue[2].ToString(), out _dateByNum))
        //        {
        //            return _dateByNum;
        //        }
        //        //If 4th Char Isn't A Slash, Day has Two Digits
        //        if (_dateCellValue[3] != '/' &&
        //            int.TryParse(_dateCellValue.Substring(2, 2), out _dateByNum))
        //        {
        //            return _dateByNum;
        //        }
        //    }
        //    else
        //    {
        //        //If 5th Char Has Slash, Day is Single Digit
        //        if (_dateCellValue[4] == '/' &&
        //            int.TryParse(_dateCellValue[3].ToString(), out _dateByNum))
        //        {
        //            return _dateByNum;
        //        }
        //        //If 5th Char Isn't A Slash, Day has Two Digits
        //        if (_dateCellValue[4] != '/' &&
        //            int.TryParse(_dateCellValue.Substring(3, 2), out _dateByNum))
        //        {
        //            return _dateByNum;
        //        }
        //    }

        //    return _dateByNum;
        //}

        //int CalculateDayFromDateSpelledOut(string _dateCellValue, int _monthCharCount)
        //{
        //    int _dateByNum = -1;
        //    int _dayFirstNumIndex = _monthCharCount + 1;
        //    int _daySecNumIndex = _monthCharCount + 2;
        //    //If 2nd Date Char Can Be Parsed, Then Date Has Two Digits
        //    if (int.TryParse(_dateCellValue[_daySecNumIndex].ToString(), out _dateByNum) &&
        //        int.TryParse(_dateCellValue.Substring(_dayFirstNumIndex, 2), out _dateByNum))
        //    {
        //        return _dateByNum;
        //    }
        //    //If 1st Char Can Be Parsed And Not The 2nd, Day has One Digit
        //    if (int.TryParse(_dateCellValue[_dayFirstNumIndex].ToString(), out _dateByNum))
        //    {
        //        return _dateByNum;
        //    }
        //    return _dateByNum;
        //}

        //EDateByMonth CalculateDateByMonth(string _dateCellValue, out bool bIsMonthSpelledOut)
        //{
        //    bIsMonthSpelledOut = false;
        //    if (DateByMonthIsSpelledOut(_dateCellValue, out var _dateByMonth))
        //    {
        //        bIsMonthSpelledOut = true;
        //        return _dateByMonth;
        //    }
        //    int _monthByNum = -1;
        //    //If 2nd Char Has Slash, Month is Single Digit
        //    if (_dateCellValue[1] == '/' &&
        //        int.TryParse(_dateCellValue[0].ToString(), out _monthByNum))
        //    {
        //        return RetrieveMonthByNumber(_monthByNum);
        //    }
        //    //If 2nd Char Isn't A Slash, Month has Two Digits
        //    if (_dateCellValue[1] != '/' &&
        //        int.TryParse(_dateCellValue.Substring(0, 2), out _monthByNum))
        //    {
        //        return RetrieveMonthByNumber(_monthByNum);
        //    }

        //    return EDateByMonth.Undecided;
        //}

        //(int dateByMonth, int dateByYear, int dateByDay) CalculateDateByMonthAndYear(string _dateCellValue)
        //{
        //    var _dateByMonth = CalculateDateByMonth(_dateCellValue, out var bIsMonthSpelledOut);
        //    (var _month, var _year) = CalculateDayMonthAYearByDateFormatted(_dateCellValue);
        //    return (_month, _year,
        //        CalculateDateByDay(_dateCellValue, bIsMonthSpelledOut, _dateByMonth));
        //}

        //public static int GetDayCountFromMonthAYear(EDateByMonth _month, int _year)
        //{
        //    switch (_month)
        //    {
        //        case EDateByMonth.Undecided:
        //            return -1;
        //        case EDateByMonth.January:
        //            return GetDayCountFromMonthAYear(1, _year);
        //        case EDateByMonth.February:
        //            return GetDayCountFromMonthAYear(2, _year);
        //        case EDateByMonth.March:
        //            return GetDayCountFromMonthAYear(3, _year);
        //        case EDateByMonth.April:
        //            return GetDayCountFromMonthAYear(4, _year);
        //        case EDateByMonth.May:
        //            return GetDayCountFromMonthAYear(5, _year);
        //        case EDateByMonth.June:
        //            return GetDayCountFromMonthAYear(6, _year);
        //        case EDateByMonth.July:
        //            return GetDayCountFromMonthAYear(7, _year);
        //        case EDateByMonth.August:
        //            return GetDayCountFromMonthAYear(8, _year);
        //        case EDateByMonth.September:
        //            return GetDayCountFromMonthAYear(9, _year);
        //        case EDateByMonth.October:
        //            return GetDayCountFromMonthAYear(10, _year);
        //        case EDateByMonth.November:
        //            return GetDayCountFromMonthAYear(11, _year);
        //        case EDateByMonth.December:
        //            return GetDayCountFromMonthAYear(12, _year);
        //        default:
        //            return -1;
        //    }
        //}

        //public static bool IsMonthMissingDays(MyMonthAYearGroupUtility _monthAYearGroup, int _dayCount)
        //{
        //    return _dayCount != GetDayCountFromMonthAYear(_monthAYearGroup.DateByMonth, _monthAYearGroup.DateByYear);
        //}
        #endregion
    }

    #region DateFormatProvider
    // Custom format provider that includes a leading zero for single-digit days
    public class MyLeadingZeroDateFormatProvider : IFormatProvider, ICustomFormatter
    {
        public object GetFormat(Type formatType)
        {
            if (formatType == typeof(ICustomFormatter))
                return this;
            return CultureInfo.CurrentCulture.GetFormat(formatType);
        }

        public string Format(string format, object arg, IFormatProvider formatProvider)
        {
            if (arg is DateTime dt)
            {
                return dt.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
            }
            return arg.ToString();
        }
    }
    #endregion

}