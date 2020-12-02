using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Corel.Interop.VGCore;

namespace ColorSpread
{
    public class BColor : ColorClass, IComparable
    {

        public int CompareTo(object obj)
        {
            if (obj == null) return 1;
            BColor bColor = obj as BColor;
            if (bColor != null)
                return this.Name.CompareTo(bColor.Name);
            else
                throw new ArgumentException("Name erro!");
        }
    }
}
