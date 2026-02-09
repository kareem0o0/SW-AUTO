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

            // --- Button Click Handler: Validate first, then create if valid ---
            btnGenerate.Click += (s, e) =>
            {
                string outFolder = @"D:\work\birr machines\parts";
                try {
                    if (!Directory.Exists(outFolder)) Directory.CreateDirectory(outFolder);
                    SldWorks swApp = (SldWorks)Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application")!);
                    swApp.Visible = true;

                    // Validate Part A
                    string validationA = ValidatePart("Part_A", aX.Text, aY.Text, aZ.Text, aHoleDia.Text, aHoleCount.Text);
                    if (!string.IsNullOrEmpty(validationA))
                    {
                        MessageBox.Show(validationA, "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return; // Stay in window, allow user to adjust
                    }

                    // Validate Part B
                    string validationB = ValidatePart("Part_B", bX.Text, bY.Text, bZ.Text, bHoleDia.Text, bHoleCount.Text);
                    if (!string.IsNullOrEmpty(validationB))
                    {
                        MessageBox.Show(validationB, "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return; // Stay in window, allow user to adjust
                    }

                    // All validations passed, now create parts
                    string partAPath = GeneratePart(swApp, "Part_A", aX.Text, aY.Text, aZ.Text, aHoleDia.Text, aHoleCount.Text, outFolder);
                    string partBPath = GeneratePart(swApp, "Part_B", bX.Text, bY.Text, bZ.Text, bHoleDia.Text, bHoleCount.Text, outFolder);

                    // Create assembly with both parts
                    GenerateAssembly(swApp, partAPath, partBPath, outFolder);

                    MessageBox.Show($"Success! Parts and Assembly saved to: {outFolder}");
                    win.Close();
                }
                catch (Exception ex) { MessageBox.Show("Error: " + ex.Message); }
            };

            win.ShowDialog();
        }

        // --- Validation Method ---
        static string ValidatePart(string name, string sx, string sy, string sz, string sDia, string sCount)
        {
            if (!double.TryParse(sx, out double x) || !double.TryParse(sy, out double y) || !double.TryParse(sz, out double z))
                return $"{name}: Invalid dimensions (Width, Height, Extrude must be numbers)";

            if (double.TryParse(sDia, out double hDia) && int.TryParse(sCount, out int hCount) && hCount > 0)
            {
                double spacingX = x / (hCount + 1);
                double radiusMM = hDia / 2;
                
                // Calculate circle positions
                var circlePositions = new System.Collections.Generic.List<(double px, double py)>();
                for (int i = 1; i <= hCount; i++)
                {
                    circlePositions.Add((spacingX * i, y / 2.0));
                }
                
                // Check for overlaps
                int overlapCount = 0;
                double totalOverlapDistance = 0;
                
                for (int i = 0; i < circlePositions.Count; i++)
                {
                    for (int j = i + 1; j < circlePositions.Count; j++)
                    {
                        double dx = circlePositions[j].px - circlePositions[i].px;
                        double dy = circlePositions[j].py - circlePositions[i].py;
                        double distance = Math.Sqrt(dx * dx + dy * dy);
                        double requiredDistance = hDia; // Minimum distance = diameter (to avoid overlap)
                        
                        if (distance < requiredDistance)
                        {
                            overlapCount++;
                            double overlapDist = requiredDistance - distance;
                            totalOverlapDistance += overlapDist;
                        }
                    }
                }
                
                // Return error if overlaps detected
                if (overlapCount > 0)
                {
                    return $"{name}: {overlapCount} circle overlap(s) detected!\n" +
                           $"Total overlap distance: {totalOverlapDistance:F2} mm\n" +
                           $"Circle diameter: {hDia} mm\n\n" +
                           $"Options to fix:\n" +
                           $"• Increase part width (X)\n" +
                           $"• Decrease hole count\n" +
                           $"• Decrease hole diameter\n\n" +
                           $"Adjust parameters and try again.";
                }
            }

            return ""; // No errors
        }
        static string GeneratePart(SldWorks swApp, string name, string sx, string sy, string sz, string sDia, string sCount, string folder)
        {
            if (!double.TryParse(sx, out double x) || !double.TryParse(sy, out double y) || !double.TryParse(sz, out double z)) return "";
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
            string fullPath = Path.Combine(folder, name + ".SLDPRT");
            swModel.SaveAs3(fullPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent);
            System.Threading.Thread.Sleep(500); // Wait for file to be written before assembly references it
            return fullPath;
        }

        static void GenerateAssembly(SldWorks swApp, string partAPath, string partBPath, string folder)
        {
            // Get assembly template
            string template = swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateAssembly);
            
            // Create new assembly document
            AssemblyDoc swAssy = (AssemblyDoc)swApp.NewDocument(template, 0, 0, 0);
            ModelDoc2 assyModel = (ModelDoc2)swAssy;

            // Insert Part A at origin
            bool resultA = swAssy.AddComponent(partAPath, 0, 0, 0);
            if (!resultA)
                throw new Exception("Failed to insert Part A into assembly");

            // Insert Part B offset in Z direction (10mm)
            bool resultB = swAssy.AddComponent(partBPath, 0, 0, 0.01); // 10mm = 0.01m
            if (!resultB)
                throw new Exception("Failed to insert Part B into assembly");

            // Rebuild assembly to make components visible
            assyModel.ForceRebuild3(false);

            // Save assembly
            string assemblyPath = Path.Combine(folder, "Final_Assembly.SLDASM");
            assyModel.SaveAs3(assemblyPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent);
        }

        static TextBox CreateInput(StackPanel p, string label, string defaultVal) {
            p.Children.Add(new TextBlock { Text = label, Margin = new Thickness(0,5,0,2) });
            TextBox t = new TextBox { Text = defaultVal, Height = 25, Background = Brushes.DimGray, Foreground = Brushes.White, BorderThickness = new Thickness(0), VerticalContentAlignment = VerticalAlignment.Center, Padding = new Thickness(5,0,0,0) };
            p.Children.Add(t);
            return t;
        }
    }
}