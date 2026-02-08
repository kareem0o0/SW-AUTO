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
            // --- UI SETUP ---
            var backColor = (Color)ColorConverter.ConvertFromString("#252526");
            var accentColor = (Color)ColorConverter.ConvertFromString("#007ACC");
            Window win = new Window { Title = "SW Smart Generator", Width = 400, Height = 450, Background = new SolidColorBrush(backColor), Foreground = Brushes.White, WindowStartupLocation = WindowStartupLocation.CenterScreen, Topmost = true };
            StackPanel stack = new StackPanel { Margin = new Thickness(20) };

            // Input Fields
            TextBox txtX = CreateInput(stack, "Base Width X (mm):");
            TextBox txtY = CreateInput(stack, "Base Height Y (mm):");
            TextBox txtZ = CreateInput(stack, "Extrude Depth (mm):");
            TextBox txtHoleDia = CreateInput(stack, "Hole Diameter (mm):");
            TextBox txtHoleCount = CreateInput(stack, "Number of Holes:");

            Button btnCreate = new Button { Content = "GENERATE WITH SAFETY CHECK", Height = 40, Margin = new Thickness(0,20,0,0), Background = new SolidColorBrush(accentColor), Foreground = Brushes.White, FontWeight = FontWeights.Bold };
            stack.Children.Add(btnCreate);
            win.Content = stack;

            btnCreate.Click += (s, e) =>
            {
                // Parse all inputs
                if (double.TryParse(txtX.Text, out double x) && double.TryParse(txtY.Text, out double y) && 
                    double.TryParse(txtZ.Text, out double z) && double.TryParse(txtHoleDia.Text, out double dia) && 
                    int.TryParse(txtHoleCount.Text, out int count))
                {
                    // --- BACKEND SAFETY & RULES ---
                    double totalHoleWidth = count * dia;
                    
                    if (dia >= y || dia >= x) {
                        MessageBox.Show($"Error: Hole Diameter ({dia}mm) is larger than the part dimensions!");
                        return;
                    }

                    if (totalHoleWidth > (x * 0.8)) { // Rule: Holes cannot take up more than 80% of the length
                        MessageBox.Show($"Error: Too many holes! {count} holes of {dia}mm exceeds safety limit for a {x}mm width. Reduce count or diameter.");
                        return;
                    }

                    btnCreate.Content = "BUILDING...";
                    RunSolidWorksTask(x, y, z, dia, count);
                    win.Close();
                }
                else { MessageBox.Show("Please enter valid numbers in all fields."); }
            };
            win.ShowDialog();
        }

        static TextBox CreateInput(StackPanel p, string label) {
            p.Children.Add(new TextBlock { Text = label, Margin = new Thickness(0,10,0,5) });
            TextBox t = new TextBox { Height = 25, Background = Brushes.DimGray, Foreground = Brushes.White, BorderThickness = new Thickness(0) };
            p.Children.Add(t);
            return t;
        }

        static void RunSolidWorksTask(double x, double y, double z, double holeDia, int holeCount)
{
    SldWorks swApp = (SldWorks)Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application")!);
    swApp.Visible = true;
    string template = swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
    ModelDoc2 swModel = (ModelDoc2)swApp.NewDocument(template, 0, 0, 0);

    // 1. START SKETCH
    swModel.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
    swModel.SketchManager.InsertSketch(true);

    // 2. DRAW RECTANGLE
    swModel.SketchManager.CreateCornerRectangle(0, 0, 0, x / 1000, y / 1000, 0);

    // 3. DRAW HOLES (While still in the same sketch)
    double centerY = (y / 2) / 1000; 
    double spacingX = (x / 1000) / (holeCount + 1);

    for (int i = 1; i <= holeCount; i++)
    {
        double currentX = spacingX * i;
        // Draw the circles inside the rectangle area
        swModel.SketchManager.CreateCircleByRadius(currentX, centerY, 0, (holeDia / 2) / 1000);
    }

    // 4. EXTRUDE EVERYTHING AT ONCE
    // Because the circles are inside the rectangle, SW treats them as voids.
    swModel.FeatureManager.FeatureExtrusion2(
        true, false, false, 
        (int)swEndConditions_e.swEndCondBlind, 
        0, z / 1000, 0, false, false, false, false, 0, 0, false, false, false, false, 
        true, true, true, 0, 0, false
    );

    swModel.ViewZoomtofit2();
}
    }
}