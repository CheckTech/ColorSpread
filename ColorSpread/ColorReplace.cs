using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Corel.Interop.VGCore;

namespace ColorSpread
{
    public class ColorReplace
    {

        private Application corelApp;
        private Color colorOrigin;
        private Color colorDestin;
        public Color ColorOrigin
        {
            get { return this.colorOrigin; }
            set
            {
                try
                {
                    this.colorOrigin = value.GetCopy();
                }
                catch (Exception erro)
                {
                    throw erro;
                }
            }
        }
        public Color ColorDestin
        {
            get { return this.colorDestin; }
            set
            {
                try
                {
                    this.colorDestin = value.GetCopy();
                }
                catch (Exception erro)
                {
                    throw erro;
                }
            }
        }


        public ColorReplace(Application corelApp)
        {
            this.corelApp = corelApp;
        }

        public void ReplaceColor(Shapes shapes)
        {

            foreach (Shape shape in shapes)
            {
                if (shape.Fill.UniformColor.IsSame(colorOrigin))
                    shape.Fill.UniformColor = colorDestin;
            }
        }
    }
}
