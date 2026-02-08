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
            // 1. Setup Colors (2026 Modern Palette)
            var backColor = (Color)ColorConverter.ConvertFromString("#252526"); // Dark Grey
            var accentColor = (Color)ColorConverter.ConvertFromString("#007ACC"); // VS Blue
            var textColor = Brushes.White;

            // 2. Create Window
            Window win = new Window
            {
                Title = "SolidWorks Geometric Creator",
                Width = 350, Height = 250,
                Background = new SolidColorBrush(backColor),
                Foreground = textColor,
                WindowStyle = WindowStyle.ThreeDBorderWindow,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                Topmost = true,
                FontFamily = new FontFamily("Segoe UI Semibold")
            };

            // 3. Layout Grid (More professional than StackPanel)
            Grid grid = new Grid { Margin = new Thickness(20) };
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

            // Labels & Inputs
            var labelX = new TextBlock { Text = "Rectangle Width (X mm):", Margin = new Thickness(0,0,0,5) };
            TextBox txtX = new TextBox { Height = 30, VerticalContentAlignment = VerticalAlignment.Center, Background = Brushes.DimGray, Foreground = Brushes.White, BorderThickness = new Thickness(0) };
            
            var labelY = new TextBlock { Text = "Rectangle Height (Y mm):", Margin = new Thickness(0,10,0,5) };
            TextBox txtY = new TextBox { Height = 30, VerticalContentAlignment = VerticalAlignment.Center, Background = Brushes.DimGray, Foreground = Brushes.White, BorderThickness = new Thickness(0) };

            // Button with "Spice" (Hover effects require XAML, but we can do solid styling here)
            Button btnCreate = new Button 
            { 
                Content = "GENERATE PART", 
                Height = 40, 
                Margin = new Thickness(0,20,0,0),
                Background = new SolidColorBrush(accentColor),
                Foreground = Brushes.White,
                FontWeight = FontWeights.Bold,
                BorderThickness = new Thickness(0)
            };

            // Add to Grid
            Grid.SetRow(labelX, 0); Grid.SetRow(txtX, 0); // Overlaying for simple example, better to use StackPanel inside Grid
            // Re-stacking for simplicity but with better spacing
            StackPanel stack = new StackPanel();
            stack.Children.Add(labelX); stack.Children.Add(txtX);
            stack.Children.Add(labelY); stack.Children.Add(txtY);
            stack.Children.Add(btnCreate);
            grid.Children.Add(stack);

            win.Content = grid;

            // 4. SolidWorks Handshake
            btnCreate.Click += (s, e) =>
            {
                if (double.TryParse(txtX.Text, out double x) && double.TryParse(txtY.Text, out double y))
                {
                    btnCreate.Content = "PROCESSING...";
                    btnCreate.IsEnabled = false;
                    RunSolidWorksTask(x, y);
                    win.Close();
                }
            };

            win.ShowDialog();
        }

        static void RunSolidWorksTask(double x, double y)
        {
            SldWorks swApp = (SldWorks)Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application")!);
            swApp.Visible = true;
            
            string template = swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
            ModelDoc2 swModel = (ModelDoc2)swApp.NewDocument(template, 0, 0, 0);

            swModel.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swModel.SketchManager.InsertSketch(true);
            
            // Convert mm to Meters
            swModel.SketchManager.CreateCornerRectangle(0, 0, 0, x/1000, y/1000, 0);
            
            swModel.SketchManager.InsertSketch(true);
            swModel.ViewZoomtofit2();
        }
    }
}