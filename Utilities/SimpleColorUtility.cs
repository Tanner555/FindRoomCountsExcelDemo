using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyCommonUtilities
{
    public class SimpleColorUtility
    {
        System.Random _myRandom;

        public SimpleColorUtility(System.Random _random = null)
        {
            _myRandom = _random != null ? _random : new System.Random();
        }

        public System.Drawing.Color GetRandomColor()
        {
            return System.Drawing.Color.FromArgb(_myRandom.Next(0, 255), _myRandom.Next(0, 255), _myRandom.Next(0, 255));
        }
    }
}
