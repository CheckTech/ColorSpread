using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Corel.Interop.VGCore;
using System.Threading.Tasks;


namespace ColorSpread
{
    //public class ValueEventArgs : EventArgs
    //{
    //    public int Value { get; set; }
    //    public ValueEventArgs(int value)
    //    {
    //        this.Value = value;
    //    }
    //}

    public enum ColorOrder
    {
        Asc,
        Desc,
        Random,
        PaletteDesc,
        PaletteAsc
    }
    public class Drawer
    {
        //public delegate void ValueEventHandler(int value);
        //public event ValueEventHandler NewTotalValueEvent;
        //public event ValueEventHandler InclementValueEvent;

        private Application corelApp;
        public Drawer(Application app)
        {
            this.corelApp = app;
            
        }
        
        public double LeftXRelativePage(ShapeRange shape, Page page)
        {
            return shape.LeftX - page.LeftX;
        }
        public double TopYRelativePage(ShapeRange shape, Page page)
        {
            return shape.TopY - page.TopY;
        }
        public double LeftX(double x, Page page)
        {
            return x + page.LeftX;
        }
        public double TopY(double y, Page page)
        {
            return y + page.TopY;
        }

        public void RandomColor(List<Color> colors, ShapeRange sr)
        {
            Random rd = new Random();
            List<int> useds = new List<int>();
            foreach (Shape item in sr)
            {



                int num = rd.Next(colors.Count);
                while (useds.Exists(r => r == num))
                {
                    num = rd.Next(colors.Count);
                    if (useds.Count == colors.Count)
                        useds.Clear();
                }
                useds.Add(num);
                item.Fill.UniformColor = colors[num];
            }
        }
        public void ColorDuplicateOrder(List<Color> colors,double margin)
        {
            this.corelApp.ActiveDocument.BeginCommandGroup("");
            this.corelApp.ActiveDocument.Unit = this.corelApp.ActiveDocument.Rulers.VUnits;
            //double margin = 10;
            bool createPage = false;
             ShapeRange shapeRange = this.corelApp.ActiveSelectionRange;
             ShapeRange pageShapes = this.corelApp.ActivePage.Shapes.All();
             pageShapes.RemoveRange(shapeRange);
            double x,y,w,h,startX,startY = 0;
            int xc = 1;
            
            int yc = 0;
            //int pageC = 1;
            startX = this.LeftXRelativePage(shapeRange, this.corelApp.ActivePage);
            startY = this.TopYRelativePage(shapeRange, this.corelApp.ActivePage);
            //shapeRange.GetPosition(out startX, out startY);
            shapeRange.GetSize(out w, out h);
           
            Page page = this.corelApp.ActiveDocument.ActivePage;


          
            for (int i = 0; i < colors.Count; i++)
            {
                shapeRange.AlignAndDistribute(cdrAlignDistributeH.cdrAlignDistributeHAlignCenter, cdrAlignDistributeV.cdrAlignDistributeVNone);
                //if (!useOrigin)
                //{
                x = startX + (w + margin) * xc;
                y = startY - (h + margin) * yc;

                    //x = ((w + margin) * xc);
                    //y = ((h + margin) * yc);
                    ////  this.corelApp.MsgShow(x.ToString());
                    if (page.SizeWidth - (x  + w) >= w)
                    {
                        xc++;

                    }
                    else
                        xc = 0;
                    if (xc == 0)
                    {
                        yc++;
                        ////this.corelApp.MsgShow(string.Format("y:{0} page:{1}",  y - startY - h,this.app.ActiveDocument.ActivePage.SizeHeight));
                        if (page.SizeHeight + (y  - h) < h )
                        {
                            createPage = true;
                            //pageC++;
                            

                            xc = 0;
                            yc = 0;
                        }
                   
                }
               
                foreach (Shape item in shapeRange.Shapes)
                {
                    if (item.Name == "color")
                        item.Fill.ApplyUniformFill(colors[i]);
                    if (item.Type == cdrShapeType.cdrTextShape)
                         item.Text.Contents = colors[i].Name;
                 
                }
                if (i == colors.Count - 1)
                    return;
               shapeRange = shapeRange.Duplicate();
               

               shapeRange.TopY = this.TopY(y, this.corelApp.ActivePage);
               shapeRange.LeftX = this.LeftX(x, this.corelApp.ActivePage);
               shapeRange.MoveToLayer(page.ActiveLayer);
              
               if (createPage)
               {
                   page = this.corelApp.ActiveDocument.InsertPages(1, false, this.corelApp.ActivePage.Index); 
                   page.Activate();

                   pageShapes.Duplicate();
                   pageShapes.MoveToLayer(page.ActiveLayer);
                   //foreach (Shape item in pageShapes)
                   //{
                   //    item.Duplicate();
                   //    item.MoveToLayer(page.ActiveLayer);
                   //}


                   createPage = false;
               }
 
                string pageName = page.Name;

            }
            this.corelApp.ActiveDocument.EndCommandGroup();
        }
        public void ChangeToBigName(List<Color> colorList,Shape activeShape)
        {
            if (activeShape.Type != cdrShapeType.cdrTextShape)
                return;
            double bigWidth = 0;
            string prevText = "";
            List<Color> temp = colorList.OrderBy(r => r.Name.Length).ToList<Color>();

            for (int i = 0; i < temp.Count; i++)
            {
                activeShape.Text.Contents = temp[i].Name;
                if(activeShape.SizeWidth > bigWidth)
                {
                    bigWidth = activeShape.SizeWidth;
                    prevText = temp[i].Name;
                }
                else
                {
                    activeShape.Text.Contents = prevText;
                }
            }

        }
        public void FlagShape(Shape activeShape)
        {
            activeShape.Name = "color";
        }
        public List<Color> ColorList(ColorOrder order,int count=-1)
        {
            List<Color> colorList = new List<Color>();
            Palette palette = this.corelApp.ActivePalette;
            if(order.Equals(ColorOrder.Random))
            {
                Random rd = new Random();
                if (palette.ColorCount == count)
                {
                     List<int> useds = new List<int>();
                     while(colorList.Count < count)
                     {



                         int num = rd.Next(1,palette.ColorCount);
                         while (useds.Exists(r => r == num))
                         {
                             num = rd.Next(1, palette.ColorCount+1);
                             if (useds.Count == count)
                                 useds.Clear();
                             
                         }
                         useds.Add(num);
                         colorList.Add(palette.Color[num]);
                     }
                     return colorList;
                }
                else
                {
                    while (colorList.Count < count)
                    {
                        int num = rd.Next(1, palette.ColorCount+1);

                        colorList.Add(palette.Color[num]);
                    }
                }
                return colorList;
            }

            
            for (int i = 1; i <= palette.ColorCount; i++)
            {
                colorList.Add(palette.Color[i] as Color);
            }
            if(order.Equals(ColorOrder.PaletteAsc))
                return colorList;
            if(order.Equals(ColorOrder.PaletteDesc))
            {
                colorList.Reverse();
                return colorList;
            }
            List<Color> colors = colorList.OrderBy(b => b.Name).ToList();
            if (order.Equals(ColorOrder.Desc))
                colors.Reverse();

           return colors;
        }

        public void ColorInWords(List<Color> colorList, bool isWords)
        {

            //Palette palette = this.corelApp.ActivePalette;

            int colorInc = 0;
            // int colorCount = palette.ColorCount;
            int colorCount = colorList.Count;
            if (isWords)
            {
                TextWords words = this.corelApp.ActiveShape.Text.Story.Words;

               // this.corelApp.ActiveDocument.BeginCommandGroup();

               var t = Task.Run(()=>{ 
                   this.corelApp.Application.Status.BeginProgress("Working");
               
                for (int i = 1; i <= words.Count; i++)
                {
                    
                    //if (InclementValueEvent != null)
                    //    InclementValueEvent(1);
                    words[i].Fill.ApplyUniformFill(colorList[colorInc]);
                    
                        colorInc++;
                   if (colorInc >= colorCount)
                        colorInc = 0;


                  this.corelApp.Application.Status.UpdateProgress();
                    
                   
                }
              //  global::System.Windows.this.corelApp.MsgShow("Test");
                this.corelApp.Application.Status.EndProgress();

               // this.corelApp.ActiveDocument.EndCommandGroup();
               });
            }
            else
            {
                TextCharacters letters = this.corelApp.ActiveShape.Text.Story.Characters;
                for (int i = 1; i <= letters.Count; i++)
                {
                    //if (InclementValueEvent != null)
                    //    InclementValueEvent(1);
                    letters[i].Fill.ApplyUniformFill(colorList[colorInc]);
                    
                        colorInc++;

                  if(colorInc>=colorCount)
                        colorInc = 0;

                }
            }
            //if (NewTotalValueEvent != null)
            //    NewTotalValueEvent(words.Count);



        }
     
     
    }
}
