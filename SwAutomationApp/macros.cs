using System;
using SolidWorks.Interop.sldworks;
using SwAutomation.Pdm;

namespace SwAutomation;

/// <summary>
/// Sample macro/orchestration flows.
///
/// The part and assembly classes know how to create themselves.
/// This file shows how those objects can be combined into larger automation scenarios.
///
/// Think of this file as the "workflow layer":
/// it decides which objects are created, how their parameters are overridden,
/// and in what order they are used together.
/// </summary>
public static class Project1
{
    /// <summary>
    /// Very small sample flow:
    /// create a few assembly documents and insert them into a machine assembly.
    /// </summary>
    public static bool savetopdm = false;
    public static void Run(string outFolder, SldWorks swApp, PdmModule pdm)
    {
        SkeletonPart skeleton = new SkeletonPart(swApp, pdm);
        skeleton.OutputFolder = outFolder;
        skeleton.SideOffset = 0.5;
        skeleton.GroundOffset = -0.25;
        skeleton.CloseAfterCreate = true;

        AssemblyFile rotor = new AssemblyFile(swApp, pdm);
        rotor.FileName = "RotorComplete.SLDASM";
        rotor.OutputFolder = outFolder;
        rotor.CloseAfterCreate = true;

        AssemblyFile stator = new AssemblyFile(swApp, pdm);
        stator.FileName = "StatorComplete.SLDASM";
        stator.OutputFolder = outFolder;
        stator.CloseAfterCreate = true;

        AssemblyFile housing = new AssemblyFile(swApp, pdm);
        housing.FileName = "HousingMachined.SLDASM";
        housing.OutputFolder = outFolder;
        housing.CloseAfterCreate = true;

        AssemblyFile machine = new AssemblyFile(swApp, pdm);
        machine.FileName = "MachineAssembly.SLDASM";
        machine.OutputFolder = outFolder;

        string skeletonPath = skeleton.Create();
        string rotorPath = rotor.Create();
        string statorPath = stator.Create();
        string housingPath = housing.Create();
        string machinePath = machine.Create();

        var inserted = machine.Insert(skeletonPath);
        machine.MateToOrigin(inserted);

        inserted = machine.Insert(rotorPath);
        machine.MateToOrigin(inserted);

        inserted = machine.Insert(statorPath);
        machine.MateToOrigin(inserted);

        inserted = machine.Insert(housingPath);
        machine.MateToOrigin(inserted);

        Console.WriteLine($"Macro completed. Machine assembly: {machinePath}");
    }

    /// <summary>
    /// Small test flow that creates skeleton and stator sheet,
    /// then mates them inside one machine assembly.
    /// </summary>
    public static void Run2(string outFolder, SldWorks swApp, PdmModule pdm)
    {
        SkeletonPart skeleton = new SkeletonPart(swApp, pdm);
        skeleton.OutputFolder = outFolder;
        skeleton.SideOffset = 2;
        skeleton.GroundOffset = -0.5;
        skeleton.CloseAfterCreate = true;

        StatorSheetPart statorSheet = new StatorSheetPart(swApp, pdm);
        statorSheet.OutputFolder = outFolder;
        statorSheet.CloseAfterCreate = true;

        AssemblyFile machine = new AssemblyFile(swApp, pdm);
        machine.FileName = "MachineAssembly.SLDASM";
        machine.OutputFolder = outFolder;

        string skeletonPath = skeleton.Create();
        string statorSheetPath = statorSheet.Create();
        string machinePath = machine.Create();

        var insertedSkeleton = machine.Insert(skeletonPath);
        machine.MateToOrigin(insertedSkeleton);
        var insertedStatorSheet = machine.Insert(statorSheetPath);
        machine.MateCoincident(insertedSkeleton, "X-Achse", insertedStatorSheet, "Z-Achse");
        machine.MateCoincident(insertedSkeleton, "Ebene rechts", insertedStatorSheet, "Ebene vorne");
        machine.MateCoincident(insertedSkeleton, "Ebene vorne", insertedStatorSheet, "Ebene rechts");

        Console.WriteLine($"Macro completed. Machine assembly: {machinePath}");
    }

    /// <summary>
    /// Reserved scratch/testing macro.
    /// </summary>
    public static void Run3(string outFolder, SldWorks swApp, PdmModule pdm)
    {
        _ = swApp;
        _ = pdm;
        Console.WriteLine("Run3 is reserved for ad-hoc generator checks.");
    }

    public static void Run4(string outFolder, SldWorks swApp, PdmModule pdm)
    {
        
        // These objects are the editable parameter surface for the machine build.
        // Skeleton planes and axes act as the assembly reference frame.
        SkeletonPart skeleton = new SkeletonPart(swApp, pdm);
        skeleton.OutputFolder = outFolder;
        skeleton.SaveToPdm = savetopdm;
        skeleton.CloseAfterCreate = true;
        skeleton.LocalFileName = "skeleton.SLDPRT";
        skeleton.SideOffset = 2;
        skeleton.GroundOffset = -0.5;
        // Optional PDM datacard example:
        // leave this only on the skeleton as a reference block.
        // If you later want the same behavior on another object, copy these lines to that object too.
        skeleton.PdmDataCard.DrawingNumber = "";
        skeleton.PdmDataCard.Title = "Skeleton";
        skeleton.PdmDataCard.Subtitle = "Automated Generation";
        skeleton.PdmDataCard.Project = "665_Birr_Project";
        skeleton.PdmDataCard.Customer = "Birr Machines AG";
        skeleton.PdmDataCard.CustomerOrder = "66";
        skeleton.PdmDataCard.Type = "44";
        skeleton.PdmDataCard.Unit = "kg";
        skeleton.PdmDataCard.CreatedFrom = "22";
        skeleton.PdmDataCard.ReplacementFor = "11";
        skeleton.PdmDataCard.DataCheck = "X";

        // Main stator sheet
        StatorSheetPart statorSheet = new StatorSheetPart(swApp, pdm);
        statorSheet.OutputFolder = outFolder;
        statorSheet.SaveToPdm = savetopdm;
        statorSheet.CloseAfterCreate = true;
        statorSheet.LocalFileName = "StatorBleche.SLDPRT";
        statorSheet.OuterDiameter = 0.99;
        statorSheet.InnerDiameter = 0.64;
        statorSheet.PlateThickness = 0.1;
        statorSheet.SlotWidth = 0.0157;
        statorSheet.SlotBottomY = 0.32;
        statorSheet.SlotTopY = 0.4052;
        statorSheet.AngleInDegrees = 70.0;
        statorSheet.FilletRadius = 0.001;
        statorSheet.SlotGuideSpacing = 0.0057;
        statorSheet.SlotGuideOffset = -0.0007;
        statorSheet.SlotPatternCount = 60;
        statorSheet.MaterialName = "AISI 1020";

        // Distance sheets 
        StatorDistanceSheetPart statorDistanceSheet = new StatorDistanceSheetPart(swApp, pdm);
        statorDistanceSheet.OutputFolder = outFolder;
        statorDistanceSheet.SaveToPdm = savetopdm;
        statorDistanceSheet.CloseAfterCreate = true;
        statorDistanceSheet.LocalFileName = "StatorDistanceBleche.SLDPRT";
        statorDistanceSheet.OuterDiameter = 0.99;
        statorDistanceSheet.InnerDiameter = 0.64;
        statorDistanceSheet.PlateThickness = 0.001;
        statorDistanceSheet.SlotWidth = 0.0205;
        statorDistanceSheet.SlotBottomY = 0.32;
        statorDistanceSheet.SlotTopY = 0.406;
        statorDistanceSheet.BossRectangleHeight = 0.16;
        statorDistanceSheet.BossRectangleWidth = 0.008;
        statorDistanceSheet.BossOuterDiameterOffset = 0.009;
        statorDistanceSheet.BossCenterlineAngleDeg = 2.96;
        statorDistanceSheet.BossExtrusionDepth = 0.01;
        statorDistanceSheet.BossCutOuterTabWidth = 0.002;
        statorDistanceSheet.BossCutOuterTabHeight = 0.0025;
        statorDistanceSheet.BossCutTopShelfThickness = 0.0015;
        statorDistanceSheet.BossCutInnerLegWidth = 0.0015;
        statorDistanceSheet.BossCutBoundaryExtension = 0.002;
        statorDistanceSheet.BossCircularPatternCount = 60;
        statorDistanceSheet.SlotPatternCount = 60;
        statorDistanceSheet.MaterialName = "AISI 1020";

        // End sheets 
        StatorEndSheetPart statorEndSheet = new StatorEndSheetPart(swApp, pdm);
        statorEndSheet.OutputFolder = outFolder;
        statorEndSheet.SaveToPdm = savetopdm;
        statorEndSheet.CloseAfterCreate = true;
        statorEndSheet.LocalFileName = "StatorEndBleche.SLDPRT";
        statorEndSheet.OuterDiameter = 0.99;
        statorEndSheet.InnerDiameter = 0.64;
        statorEndSheet.PlateThickness = 0.001;
        statorEndSheet.SlotWidth = 0.0205;
        statorEndSheet.SlotBottomY = 0.32;
        statorEndSheet.SlotTopY = 0.406;
        statorEndSheet.SlotPatternCount = 60;
        statorEndSheet.MaterialName = "AISI 1020";

        // Torsion bars 
        TorsionBarPart torsionBar = new TorsionBarPart(swApp, pdm);
        torsionBar.OutputFolder = outFolder;
        torsionBar.SaveToPdm = savetopdm;
        torsionBar.CloseAfterCreate = true;
        torsionBar.LocalFileName = "TorsionBar.SLDPRT";
        torsionBar.BarLength = 1.074;
        torsionBar.BarHeight = 0.04;
        torsionBar.BarThickness = 0.03;
        torsionBar.HoleCenterlineOffsetFromBottom = 0.02;
        torsionBar.OuterHoleEndOffset = 0.03;
        torsionBar.HolePairSpacing = 0.315;
        torsionBar.OuterHoleDiameter = 0.01;
        torsionBar.InnerHoleDiameter = 0.016;
        torsionBar.CenterHoleDiameter = 0.016;
        torsionBar.OuterTapSizePrimary = "M10x1.5";
        torsionBar.OuterTapSizeFallback = "M10";
        torsionBar.InnerTapSizePrimary = "M16x2";
        torsionBar.InnerTapSizeFallback = "M16";
        torsionBar.P0001ConfigName = "P0001";
        torsionBar.P0002ConfigName = "P0002";
        torsionBar.MaterialName = "AISI 1020";

        // Press plates 
        PressPlatePart pressPlate = new PressPlatePart(swApp, pdm);
        pressPlate.OutputFolder = outFolder;
        pressPlate.SaveToPdm = savetopdm;
        pressPlate.CloseAfterCreate = true;
        pressPlate.LocalFileName = "PressPlate.SLDPRT";
        pressPlate.OuterDiameter = 0.99;
        pressPlate.RingInnerDiameter = 0.84;
        pressPlate.PlateOuterInsetFromOuterDiameter = 0.005;
        pressPlate.PlateRadialLength = 0.165;
        pressPlate.RingThickness = 0.002;
        pressPlate.PlateBodyThickness = 0.01;
        pressPlate.PlateWidth = 0.006;
        pressPlate.PlateCount = 60;
        pressPlate.AssemblyAngleDeg = 3.0;
        pressPlate.MaterialName = "AISI 1020";

        // The NDE press ring 
        StatorPressringNdePart pressRingNde = new StatorPressringNdePart(swApp, pdm);
        pressRingNde.OutputFolder = outFolder;
        pressRingNde.SaveToPdm = savetopdm;
        pressRingNde.CloseAfterCreate = true;
        pressRingNde.LocalFileName = "StatorPressringNDE.SLDPRT";
        pressRingNde.OuterDiameter = 1.1;
        pressRingNde.InnerDiameter = 0.84;
        pressRingNde.PressRingOuterDiameter = 0.86;
        pressRingNde.RingThickness = 0.028;
        pressRingNde.PressRingThickness = 0.002;
        pressRingNde.BaseInnerChamferDistance = 0.02;
        pressRingNde.BaseInnerChamferAngleDeg = 30.0;
        pressRingNde.PocketCenterRadius = 0.52;
        pressRingNde.PocketWidth = 0.043;
        pressRingNde.PocketHeight = 0.036;
        pressRingNde.PocketCornerRadius = 0.005;
        pressRingNde.PocketCount = 8;
        pressRingNde.MaterialName = "AISI 1020";

        // This is the target assembly document that all generated parts are inserted into.
        AssemblyFile machine = new AssemblyFile(swApp, pdm);
        machine.OutputFolder = outFolder;
        machine.SaveToPdm = savetopdm;
        machine.FileName = "MachineAssembly.SLDASM";
        machine.CloseAfterCreate = false;
        int repeatedDistanceEndSheetPacks = 5;

        // These values are derived from the editable object settings above.
        // If you change part thicknesses or pocket counts, the assembly spacing below follows automatically.
        double statorSheetPackThickness = statorSheet.PlateThickness;
        double statorDistanceSheetStackThickness = statorDistanceSheet.PlateThickness + statorDistanceSheet.BossExtrusionDepth;
        double statorEndSheetStackThickness = statorEndSheet.PlateThickness;
        double pressPlateStackThickness = Math.Max(pressPlate.PlateBodyThickness, pressPlate.RingThickness);
        double pressRingNdeStackThickness = pressRingNde.RingThickness;
        double repeatedStackBlockThickness = statorDistanceSheetStackThickness + statorEndSheetStackThickness + statorSheetPackThickness;
        double torsionBarSlotRadius = pressRingNde.PocketCenterRadius;
        int torsionBarPatternCount = pressRingNde.PocketCount;

        // Create each source file first, then start inserting those saved files into the assembly.
        string skeletonPath = skeleton.Create();
        string statorSheetPath = statorSheet.Create();
        string statorDistanceSheetPath = statorDistanceSheet.Create();
        string statorEndSheetPath = statorEndSheet.Create();
        string torsionBarPath = torsionBar.CreatePart();
        string pressPlatePath = pressPlate.Create();
        string pressRingNdePath = pressRingNde.Create();
        string machinePath = machine.Create();

        // Start from the skeleton and build the axial stack from one end.
        Component2 insertedSkeleton = machine.Insert(skeletonPath);
        machine.MateToOrigin(insertedSkeleton);

        double stackOffset = 0.0;
        Component2 insertedStackComponent;

        insertedStackComponent = machine.Insert(pressRingNdePath);
        machine.MateCoincident(insertedSkeleton, "X-Achse", insertedStackComponent, "Z-Achse");
        machine.MateCoincident(insertedSkeleton, "Ebene vorne", insertedStackComponent, "Ebene rechts");
        machine.MateCoincident(insertedSkeleton, "Ebene rechts", insertedStackComponent, "Ebene vorne", stackOffset);
        stackOffset += pressRingNdeStackThickness;

        insertedStackComponent = machine.Insert(pressPlatePath);
        machine.MateCoincident(insertedSkeleton, "X-Achse", insertedStackComponent, "Z-Achse");
        machine.MateAngle(insertedSkeleton, "Ebene vorne", insertedStackComponent, "Ebene rechts", pressPlate.AssemblyAngleDeg);
        machine.MateCoincident(insertedSkeleton, "Ebene rechts", insertedStackComponent, "Ebene vorne", stackOffset);
        stackOffset += pressPlateStackThickness;

        insertedStackComponent = machine.Insert(statorEndSheetPath);
        machine.MateCoincident(insertedSkeleton, "X-Achse", insertedStackComponent, "Z-Achse");
        machine.MateCoincident(insertedSkeleton, "Ebene vorne", insertedStackComponent, "Ebene rechts");
        machine.MateCoincident(insertedSkeleton, "Ebene rechts", insertedStackComponent, "Ebene vorne", stackOffset);
        stackOffset += statorEndSheetStackThickness;

        insertedStackComponent = machine.Insert(statorSheetPath);
        machine.MateCoincident(insertedSkeleton, "X-Achse", insertedStackComponent, "Z-Achse");
        machine.MateCoincident(insertedSkeleton, "Ebene vorne", insertedStackComponent, "Ebene rechts");
        machine.MateCoincident(insertedSkeleton, "Ebene rechts", insertedStackComponent, "Ebene vorne", stackOffset);
        stackOffset += statorSheetPackThickness;

        Component2 repeatedDistanceSeed = machine.Insert(statorDistanceSheetPath);
        machine.MateCoincident(insertedSkeleton, "X-Achse", repeatedDistanceSeed, "Z-Achse");
        machine.MateCoincident(insertedSkeleton, "Ebene vorne", repeatedDistanceSeed, "Ebene rechts");
        machine.MateCoincident(insertedSkeleton, "Ebene rechts", repeatedDistanceSeed, "Ebene vorne", stackOffset);
        stackOffset += statorDistanceSheetStackThickness;

        Component2 repeatedEndSeed = machine.Insert(statorEndSheetPath);
        machine.MateCoincident(insertedSkeleton, "X-Achse", repeatedEndSeed, "Z-Achse");
        machine.MateCoincident(insertedSkeleton, "Ebene vorne", repeatedEndSeed, "Ebene rechts");
        machine.MateCoincident(insertedSkeleton, "Ebene rechts", repeatedEndSeed, "Ebene vorne", stackOffset);
        stackOffset += statorEndSheetStackThickness;

        Component2 repeatedStatorSeed = machine.Insert(statorSheetPath);
        machine.MateCoincident(insertedSkeleton, "X-Achse", repeatedStatorSeed, "Z-Achse");
        machine.MateCoincident(insertedSkeleton, "Ebene vorne", repeatedStatorSeed, "Ebene rechts");
        machine.MateCoincident(insertedSkeleton, "Ebene rechts", repeatedStatorSeed, "Ebene vorne", stackOffset);
        stackOffset += statorSheetPackThickness;

        // Repeat the middle stator block along the machine length.
        machine.LinearPattern(
            insertedSkeleton,
            "X-Achse",
            repeatedDistanceEndSheetPacks + 1,
            repeatedStackBlockThickness,
            repeatedDistanceSeed,
            repeatedEndSeed,
            repeatedStatorSeed);

        // Close the far end with the same hardware sequence used at the start.
        stackOffset += repeatedDistanceEndSheetPacks * repeatedStackBlockThickness;

        insertedStackComponent = machine.Insert(statorEndSheetPath);
        machine.MateCoincident(insertedSkeleton, "X-Achse", insertedStackComponent, "Z-Achse");
        machine.MateCoincident(insertedSkeleton, "Ebene vorne", insertedStackComponent, "Ebene rechts", 0, true);
        machine.MateCoincident(insertedSkeleton, "Ebene rechts", insertedStackComponent, "Ebene vorne", stackOffset);
        stackOffset += statorEndSheetStackThickness;

        insertedStackComponent = machine.Insert(pressPlatePath);
        machine.MateCoincident(insertedSkeleton, "X-Achse", insertedStackComponent, "Z-Achse");
        machine.MateAngle(insertedSkeleton, "Ebene vorne", insertedStackComponent, "Ebene rechts", pressPlate.AssemblyAngleDeg);
        machine.MateCoincident(insertedSkeleton, "Ebene rechts", insertedStackComponent, "Ebene vorne", stackOffset);
        stackOffset += pressPlateStackThickness;

        insertedStackComponent = machine.Insert(pressRingNdePath);
        machine.MateCoincident(insertedSkeleton, "X-Achse", insertedStackComponent, "Z-Achse");
        machine.MateCoincident(insertedSkeleton, "Ebene vorne", insertedStackComponent, "Ebene rechts");
        machine.MateCoincident(insertedSkeleton, "Ebene rechts", insertedStackComponent, "Ebene vorne", stackOffset);
        stackOffset += pressRingNdeStackThickness;

        // Place one torsion bar and pattern it around the main axis.
        Component2 insertedTorsionBar = machine.Insert(torsionBarPath);
        machine.MateParallel(insertedSkeleton, "X-Achse", insertedTorsionBar, "X-Achse");
        machine.MateCoincident(insertedSkeleton, "Ebene vorne", insertedTorsionBar, "Ebene oben");
        machine.MateCoincident(insertedSkeleton, "Ebene oben", insertedTorsionBar, "Ebene vorne", torsionBarSlotRadius);
        machine.MateCoincident(insertedSkeleton, "Ebene rechts", insertedTorsionBar, "Ebene rechts", stackOffset / 2.0);
        machine.CircularPattern(insertedSkeleton, "X-Achse", torsionBarPatternCount, 2 * Math.PI, insertedTorsionBar);

        // Finalize the assembly only after the full build is done.
        // Local workflow:
        // - CloseAfterCreate = true -> save and close
        // - CloseAfterCreate = false -> save and leave open
        // PDM workflow:
        // - always close here so the file can be saved to PDM and its card can be filled
        if (machine.SaveToPdm || machine.CloseAfterCreate)
        {
            machine.Close();
        }
        else
        {
            machine.Save();
        }

        Console.WriteLine($"Macro completed. Machine assembly: {machinePath}");
    }

    public static void Run5(string outFolder, SldWorks swApp, PdmModule pdm)
    {
        // This is the object you edit for the drawing test.
        // It owns everything for the torsion bar:
        // 1. the 3D model parameters
        // 2. the 2D drawing parameters
        TorsionBarPart torsionBar = new TorsionBarPart(swApp, pdm);
        torsionBar.OutputFolder = outFolder;
        torsionBar.SaveToPdm = savetopdm;
        torsionBar.CloseAfterCreate = true;
        torsionBar.LocalFileName = "TorsionBar.SLDPRT";
        torsionBar.BarLength = 1.074;
        torsionBar.BarHeight = 0.04;
        torsionBar.BarThickness = 0.03;
        torsionBar.HoleCenterlineOffsetFromBottom = 0.02;
        torsionBar.OuterHoleEndOffset = 0.03;
        torsionBar.HolePairSpacing = 0.315;
        torsionBar.OuterHoleDiameter = 0.01;
        torsionBar.InnerHoleDiameter = 0.016;
        torsionBar.CenterHoleDiameter = 0.016;
        torsionBar.OuterTapSizePrimary = "M10x1.5";
        torsionBar.OuterTapSizeFallback = "M10";
        torsionBar.InnerTapSizePrimary = "M16x2";
        torsionBar.InnerTapSizeFallback = "M16";
        torsionBar.P0001ConfigName = "P0001";
        torsionBar.P0002ConfigName = "P0002";
        torsionBar.MaterialName = "AISI 1020";

        // These are drawing-only settings.
        // They live on the same part object so you can configure everything in one place,
        // but the actual drawing work is still implemented in drawing.cs.
        torsionBar.DrawingOutputFolder = outFolder;
        torsionBar.DrawingSaveToPdm = savetopdm;
        // false = leave the drawing open after Create(); true = close it after saving
        torsionBar.DrawingCloseAfterCreate = false;
        torsionBar.DrawingLocalFileName = "TorsionBar.SLDDRW";
        torsionBar.DrawingSheetName = "Torsion Bar";
        torsionBar.DrawingLanguageCode = "EN";
        torsionBar.DrawingTemplateFolderPath = @"C:\Users\kareem.salah\PDM\Birr Machines PDM\40_Templates\Solidworks\Blattformate\Birr Machines";
        torsionBar.DrawingBottomTitleBlockClearance = 0.085;
        torsionBar.DrawingReferencedConfiguration = "P0002";

        // Call the combined entry point on the part.
        // Internally:
        // torsionBar.Create()
        // -> creates the 3D part by calling CreatePart()
        // -> calls the torsion-bar drawing method in drawing.cs
        // -> passes this torsionBar object into that method
        // -> drawing.cs generates the drawing from the current part data
        string drawingPath = torsionBar.Create();
        Console.WriteLine($"Run5 completed. Drawing: {drawingPath}");
    }
}



