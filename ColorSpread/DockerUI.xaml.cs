using br.corp.bonus630.plugin.ZxingQrCodeConfigurator;
using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;


namespace ColorSpread
{

    public partial class DockerUI : UserControl
    {
        Corel.Interop.VGCore.Application corelApp;
        private Drawer drawer;
        private ColorReplace colorReplace;
        private List<Corel.Interop.VGCore.Color> colorList;
        private Styles.StylesController stylesController;
        public DockerUI(Corel.Interop.VGCore.Application app)
        {
            InitializeComponent();
            drawer = new Drawer(app);
            colorReplace = new ColorReplace(app);
            this.corelApp = app;
            stylesController = new Styles.StylesController(this.Resources, app);
            //this.corelApp.OpenDocument(@"CUsersReginaldoDesktop\cores.cdr");
            //drawer.NewTotalValueEvent += (v) => { this.Dispatcher.Invoke(new Action(() => { progressBar.Value = 0; progressBar.Maximum = v; })); };
            //drawer.InclementValueEvent += (v) => { this.Dispatcher.Invoke(new Action(() => { progressBar.Value++; })); };
        }
        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            stylesController.LoadThemeFromPreference();
        }
        private void btn_show_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            ActionRunner(() =>
            {
                colorReplace.ReplaceColor(this.corelApp.ActiveLayer.Shapes);
            });
        }

        private void img_color_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            ActionRunner(() =>
            {
            //double x = 0;
            //double y = 0;
            //int shift = 0;
            //this.corelApp.ActiveDocument.GetUserClick(out x, out y, out shift, 1, false, Corel.Interop.VGCore.cdrCursorShape.cdrCursorExtPick);
            //Corel.Interop.VGCore.Shapes selectedShapes = this.corelApp.ActiveDocument.ActivePage.SelectShapesAtPoint(x, y, false).Shapes;
            //if (selectedShapes == null || selectedShapes.Count < 1)
            //    return;
            //colorReplace.ColorOrigin = selectedShapes.First.Fill.UniformColor;
            ColorPicker c = new ColorPicker(this.corelApp.ActivePalette);
                if ((bool)c.ShowDialog())
                {
                    colorReplace.ColorOrigin = c.SelectedColor.CorelColor;
                    img_color.Background = colorReplace.ColorOrigin.ToSystemColor();
                }
            });
        }

        private void img_color2_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            ActionRunner(() =>
            {
                //colorReplace.ColorDestin = this.corelApp.ActivePalette.Color[20];
                //double x = 0;
                //double y = 0;
                //int shift = 0;
                //this.corelApp.ActiveDocument.GetUserClick(out x, out y, out shift, 1, false, Corel.Interop.VGCore.cdrCursorShape.cdrCursorExtPick);
                //Corel.Interop.VGCore.Shapes selectedShapes = this.corelApp.ActiveDocument.ActivePage.SelectShapesAtPoint(x, y, false).Shapes;
                //if (selectedShapes == null || selectedShapes.Count < 1)
                //    return;
                ColorPicker c = new ColorPicker(this.corelApp.ActivePalette);
                if ((bool)c.ShowDialog())
                {
                    colorReplace.ColorDestin = c.SelectedColor.CorelColor;
                    img_color.Background = colorReplace.ColorOrigin.ToSystemColor();
                }
               // colorReplace.ColorDestin = selectedShapes.First.Fill.UniformColor;
               // img_color2.Background = colorReplace.ColorDestin.ToSystemColor();
            });
        }

        private void btn_duplicateOrder_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            ActionRunner(() =>
            {

                fillColorList(this.corelApp.ActivePalette.ColorCount);
                drawer.ColorDuplicateOrder(this.colorList, slider_margin.Value);
            });
        }



        private void btn_colorinWords_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            ActionRunner(() =>
            {

                fillColorList(this.corelApp.ActivePalette.ColorCount);
                drawer.ColorInWords(this.colorList, (bool)rb_words.IsChecked);
            });
        }



        private void btn_random_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            this.corelApp.Optimization = true;
            try
            {
                fillColorList(this.corelApp.ActiveSelectionRange.Shapes.Count);
                drawer.RandomColor(this.colorList, this.corelApp.ActiveDocument.SelectionRange);
            }

            catch
            {
                this.corelApp.MsgShow("Select many shapes to apply");
            }
            finally
            {
                this.corelApp.Optimization = false;
                this.corelApp.Refresh();
            }
        }
        private void fillColorList(int count = -1)
        {
            if ((bool)rb_asc.IsChecked)
                this.colorList = drawer.ColorList(ColorOrder.Asc);
            if ((bool)rb_desc.IsChecked)
                this.colorList = drawer.ColorList(ColorOrder.Desc);
            if ((bool)rb_palette_desc.IsChecked)
                this.colorList = drawer.ColorList(ColorOrder.PaletteDesc);
            if ((bool)rb_palette_asc.IsChecked)
                this.colorList = drawer.ColorList(ColorOrder.PaletteAsc);
            if ((bool)rb_random.IsChecked)
                this.colorList = drawer.ColorList(ColorOrder.Random, count);

        }

        private void btn_position_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            ActionRunner(() =>
            {
                double x = 0;
                double y = 0;
                //this.corelApp.ActiveShape.GetPosition(out x, out y);
                x = drawer.LeftXRelativePage(this.corelApp.ActiveSelectionRange, this.corelApp.ActivePage);
                y = drawer.TopYRelativePage(this.corelApp.ActiveSelectionRange, this.corelApp.ActivePage);
                this.corelApp.MsgShow(string.Format("X:{0} Y:{1}", x, y));
            });
        }

        private void Slider_ValueChanged(object sender, System.Windows.RoutedPropertyChangedEventArgs<double> e)
        {
            lba_margin.Content = e.NewValue.ToString();
        }

        private void btn_bigColorName_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            this.corelApp.Optimization = true;
            fillColorList(this.corelApp.ActiveSelectionRange.Shapes.Count);
            try
            {
                drawer.ChangeToBigName(this.colorList, this.corelApp.ActiveDocument.ActiveShape);
            }

            catch
            {
                this.corelApp.MsgShow("Mark a text as a template for repetition");
            }
            this.corelApp.Optimization = false;
            this.corelApp.Refresh();
        }

        private void Button_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            try
            {
                drawer.FlagShape(this.corelApp.ActiveDocument.ActiveShape);
            }
            catch
            {
                this.corelApp.MsgShow("Mark a shape as a template for repetition");
            }
        }
        private void ActionRunner(Action action)
        {

            this.corelApp.Optimization = true;
            this.corelApp.EventsEnabled = false;
            try
            {
                action.Invoke();
            }

            catch
            {

            }
            finally
            {
                this.corelApp.Optimization = false;
                this.corelApp.EventsEnabled = true;
                this.corelApp.Refresh();
            }
        }

    }
}
