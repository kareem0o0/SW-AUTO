using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SwAutomation
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            // 1. Setup Colors
            var backColor = (Color)ColorConverter.ConvertFromString("#252526");
            var accentColor = (Color)ColorConverter.ConvertFromString("#007ACC");
            var textColor = Brushes.White;

            // 2. Create Window (Increased height to 320 to fit the extra field)
            Window win = new Window
            {
                Title = "SolidWorks Geometric Creator",
                Width = 350, Height = 320, 
                Background = new SolidColorBrush(backColor),
                Foreground = textColor,
                WindowStyle = WindowStyle.ThreeDBorderWindow,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                Topmost = true,
                FontFamily = new FontFamily("Segoe UI Semibold")
            };

            Grid grid = new Grid { Margin = new Thickness(20) };
            StackPanel stack = new StackPanel();

            // Inputs
            var labelX = new TextBlock { Text = "Width X (mm):", Margin = new Thickness(0,0,0,5) };
            TextBox txtX = new TextBox { Height = 30, Background = Brushes.DimGray, Foreground = Brushes.White, BorderThickness = new Thickness(0), VerticalContentAlignment = VerticalAlignment.Center };
            
            var labelY = new TextBlock { Text = "Height Y (mm):", Margin = new Thickness(0,10,0,5) };
            TextBox txtY = new TextBox { Height = 30, Background = Brushes.DimGray, Foreground = Brushes.White, BorderThickness = new Thickness(0), VerticalContentAlignment = VerticalAlignment.Center };

            // NEW: Extrusion Depth Input
            var labelZ = new TextBlock { Text = "Extrusion Depth (mm):", Margin = new Thickness(0,10,0,5) };
            TextBox txtZ = new TextBox { Height = 30, Background = Brushes.DimGray, Foreground = Brushes.White, BorderThickness = new Thickness(0), VerticalContentAlignment = VerticalAlignment.Center };

            Button btnCreate = new Button 
            { 
                Content = "GENERATE 3D PART", 
                Height = 40, 
                Margin = new Thickness(0,20,0,0),
                Background = new SolidColorBrush(accentColor),
                Foreground = Brushes.White,
                FontWeight = FontWeights.Bold,
                BorderThickness = new Thickness(0)
            };

            // Assembly
            stack.Children.Add(labelX); stack.Children.Add(txtX);
            stack.Children.Add(labelY); stack.Children.Add(txtY);
            stack.Children.Add(labelZ); stack.Children.Add(txtZ);
            stack.Children.Add(btnCreate);
            grid.Children.Add(stack);
            win.Content = grid;

            // 4. SolidWorks Handshake
            btnCreate.Click += (s, e) =>
            {
                if (double.TryParse(txtX.Text, out double x) && 
                    double.TryParse(txtY.Text, out double y) && 
                    double.TryParse(txtZ.Text, out double z))
                {
                    btnCreate.Content = "EXTRUDING...";
                    btnCreate.IsEnabled = false;
                    RunSolidWorksTask(x, y, z);
                    win.Close();
                }
                else
                {
                    MessageBox.Show("Please enter valid numeric values for all dimensions.");
                }
            };

            win.ShowDialog();
        }

        static void RunSolidWorksTask(double x, double y, double z)
{
    SldWorks swApp = (SldWorks)Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application")!);
    swApp.Visible = true;
    
    string template = swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
    ModelDoc2 swModel = (ModelDoc2)swApp.NewDocument(template, 0, 0, 0);

    // 1. Create Sketch
    swModel.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
    swModel.SketchManager.InsertSketch(true);
    swModel.SketchManager.CreateCornerRectangle(0, 0, 0, x/1000, y/1000, 0);
    
    // 2. Extrude Boss (Updated with all 23 required parameters)
    swModel.FeatureManager.FeatureExtrusion2(
        true,               // 1. Sd (Single Direction)
        false,              // 2. Flip
        false,              // 3. Dir
        (int)swEndConditions_e.swEndCondBlind, // 4. T1 (End Condition)
        0,                  // 5. T2
        z / 1000,           // 6. D1 (Depth in Meters)
        0,                  // 7. D2
        false,              // 8. Dchk1
        false,              // 9. Dchk2
        false,              // 10. Ddir1
        false,              // 11. Ddir2
        0,                  // 12. Dang1
        0,                  // 13. Dang2
        false,              // 14. OffsetReverse1
        false,              // 15. OffsetReverse2
        false,              // 16. TranslateSurface1
        false,              // 17. TranslateSurface2
        true,               // 18. Merge
        true,               // 19. UseFeatScope
        true,               // 20. UseAutoSelect
        0,                  // 21. T0
        0,                  // 22. StartOffset
        false               // 23. FlipStartOffset
    );

    swModel.ViewZoomtofit2();
}
    }
}