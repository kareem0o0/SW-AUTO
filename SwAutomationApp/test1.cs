using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging; // Required for BitmapImage
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SwAutomation
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            var backColor = (Color)ColorConverter.ConvertFromString("#252526");
            var accentColor = (Color)ColorConverter.ConvertFromString("#007ACC");

            Window win = new Window { 
                Title = "SolidWorks Multi-Part Dashboard with Reference", 
                Width = 1000, Height = 550, // Widened to fit the image column
                Background = new SolidColorBrush(backColor), 
                Foreground = Brushes.White, 
                WindowStartupLocation = WindowStartupLocation.CenterScreen 
            };

            StackPanel mainLayout = new StackPanel { Margin = new Thickness(20) };

            // 1. Grid with THREE Columns
            Grid columnGrid = new Grid();
            columnGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) }); // Part A
            columnGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) }); // Part B
            columnGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1.5, GridUnitType.Star) }); // Image (slightly wider)

            // --- LEFT COLUMN (PART A) ---
            StackPanel leftStack = new StackPanel { Margin = new Thickness(10) };
            leftStack.Children.Add(new TextBlock { Text = "PART A", FontSize = 18, FontWeight = FontWeights.Bold, Foreground = new SolidColorBrush(accentColor), Margin = new Thickness(0,0,0,10) });
            TextBox aX = CreateInput(leftStack, "Width X (mm):", "100");
            TextBox aY = CreateInput(leftStack, "Height Y (mm):", "50");
            TextBox aZ = CreateInput(leftStack, "Extrude Z (mm):", "10");
            TextBox aHoleDia = CreateInput(leftStack, "Hole Dia (mm):", "");
            TextBox aHoleCount = CreateInput(leftStack, "Hole Count:", "");
            Grid.SetColumn(leftStack, 0);
            columnGrid.Children.Add(leftStack);

            // --- MIDDLE COLUMN (PART B) ---
            StackPanel midStack = new StackPanel { Margin = new Thickness(10) };
            midStack.Children.Add(new TextBlock { Text = "PART B", FontSize = 18, FontWeight = FontWeights.Bold, Foreground = new SolidColorBrush(accentColor), Margin = new Thickness(0,0,0,10) });
            TextBox bX = CreateInput(midStack, "Width X (mm):", "80");
            TextBox bY = CreateInput(midStack, "Height Y (mm):", "80");
            TextBox bZ = CreateInput(midStack, "Extrude Z (mm):", "20");
            TextBox bHoleDia = CreateInput(midStack, "Hole Dia (mm):", "");
            TextBox bHoleCount = CreateInput(midStack, "Hole Count:", "");
            Grid.SetColumn(midStack, 1);
            columnGrid.Children.Add(midStack);

            // --- RIGHT COLUMN (IMAGE REFERENCE) ---
            StackPanel imageStack = new StackPanel { Margin = new Thickness(10) };
            imageStack.Children.Add(new TextBlock { Text = "REFERENCE MODEL", FontSize = 18, FontWeight = FontWeights.Bold, Foreground = new SolidColorBrush(accentColor), Margin = new Thickness(0,0,0,10) });
            
            Image refImage = new Image { 
                Stretch = Stretch.Uniform, 
                Margin = new Thickness(0, 5, 0, 0),
                HorizontalAlignment = HorizontalAlignment.Center
            };

            // Loading the image from your path
            string imagePath = @"D:\work\birr machines\assets\ref1.jpg";
            if (File.Exists(imagePath))
            {
                refImage.Source = new BitmapImage(new Uri(imagePath));
            }
            else
            {
                imageStack.Children.Add(new TextBlock { Text = "Image not found at path:\n" + imagePath, Foreground = Brushes.Red, TextWrapping = TextWrapping.Wrap });
            }

            imageStack.Children.Add(refImage);
            Grid.SetColumn(imageStack, 2);
            columnGrid.Children.Add(imageStack);

            // 2. Generate Button
            Button btnGenerate = new Button { 
                Content = "GENERATE & SAVE TO D:\\WORK\\...", 
                Height = 50, Margin = new Thickness(10,30,10,0), 
                Background = new SolidColorBrush(accentColor), 
                Foreground = Brushes.White, FontWeight = FontWeights.Bold 
            };

            mainLayout.Children.Add(columnGrid);
            mainLayout.Children.Add(btnGenerate);
            win.Content = mainLayout;

            // --- REUSE YOUR EXISTING GENERATEPART LOGIC HERE ---
            btnGenerate.Click += (s, e) =>
            {
                string outFolder = @"D:\work\birr machines\parts";
                try {
                    if (!Directory.Exists(outFolder)) Directory.CreateDirectory(outFolder);
                    SldWorks swApp = (SldWorks)Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application")!);
                    swApp.Visible = true;

                    GeneratePart(swApp, "Part_A", aX.Text, aY.Text, aZ.Text, aHoleDia.Text, aHoleCount.Text, outFolder);
                    GeneratePart(swApp, "Part_B", bX.Text, bY.Text, bZ.Text, bHoleDia.Text, bHoleCount.Text, outFolder);

                    MessageBox.Show($"Success! Parts saved to: {outFolder}");
                    win.Close();
                }
                catch (Exception ex) { MessageBox.Show("Error: " + ex.Message); }
            };

            win.ShowDialog();
        }

        // ... (Keep your GeneratePart and CreateInput methods the same as before)
        static void GeneratePart(SldWorks swApp, string name, string sx, string sy, string sz, string sDia, string sCount, string folder)
        {
            if (!double.TryParse(sx, out double x) || !double.TryParse(sy, out double y) || !double.TryParse(sz, out double z)) return;
            string template = swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
            ModelDoc2 swModel = (ModelDoc2)swApp.NewDocument(template, 0, 0, 0);
            swModel.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swModel.SketchManager.InsertSketch(true);
            swModel.SketchManager.CreateCornerRectangle(0, 0, 0, x / 1000, y / 1000, 0);
            if (double.TryParse(sDia, out double hDia) && int.TryParse(sCount, out int hCount) && hCount > 0)
            {
                double spacingX = (x / 1000) / (hCount + 1);
                for (int i = 1; i <= hCount; i++)
                    swModel.SketchManager.CreateCircleByRadius(spacingX * i, (y / 2000), 0, (hDia / 2000));
            }
            swModel.FeatureManager.FeatureExtrusion2(true, false, false, (int)swEndConditions_e.swEndCondBlind, 0, z / 1000, 0, false, false, false, false, 0, 0, false, false, false, false, true, true, true, 0, 0, false);
            swModel.SaveAs3(Path.Combine(folder, name + ".SLDPRT"), (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent);
        }

        static TextBox CreateInput(StackPanel p, string label, string defaultVal) {
            p.Children.Add(new TextBlock { Text = label, Margin = new Thickness(0,5,0,2) });
            TextBox t = new TextBox { Text = defaultVal, Height = 25, Background = Brushes.DimGray, Foreground = Brushes.White, BorderThickness = new Thickness(0), VerticalContentAlignment = VerticalAlignment.Center, Padding = new Thickness(5,0,0,0) };
            p.Children.Add(t);
            return t;
        }
    }
}