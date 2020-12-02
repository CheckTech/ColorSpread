using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Corel.Interop.VGCore;

namespace ColorSpread
{
    public static class MethodsExtension
    {
        public static  System.Windows.Media.SolidColorBrush ToSystemColor(this Color corelColor)
        {
            string hexValue = corelColor.HexValue;
            System.Windows.Media.Color color = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(hexValue);
            System.Windows.Media.SolidColorBrush b = new System.Windows.Media.SolidColorBrush(color);
            return b;
        }
       
       
    }
}
