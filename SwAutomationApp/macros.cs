using System;
using SolidWorks.Interop.sldworks;

namespace SwAutomation;

public static class Project1
{
    public static void Run(string outFolder, SldWorks swApp)
    {
        if (string.IsNullOrWhiteSpace(outFolder))
            throw new ArgumentException("Output folder is required.", nameof(outFolder));

        SkeletonPart skeleton = new SkeletonPart(swApp);
        skeleton.OutputFolder = outFolder;
        skeleton.DESideOffset = 0.5;
        skeleton.NDESideOffset = 0.5;
        skeleton.GroundOffsetMm = -0.25;
        skeleton.CloseAfterCreate = true;

        AssemblyFile rotor = new AssemblyFile(swApp);
        rotor.FileName = "RotorComplete.SLDASM";
        rotor.OutputFolder = outFolder;
        rotor.CloseAfterCreate = true;

        AssemblyFile stator = new AssemblyFile(swApp);
        stator.FileName = "StatorComplete.SLDASM";
        stator.OutputFolder = outFolder;
        stator.CloseAfterCreate = true;

        AssemblyFile housing = new AssemblyFile(swApp);
        housing.FileName = "HousingMachined.SLDASM";
        housing.OutputFolder = outFolder;
        housing.CloseAfterCreate = true;

        AssemblyFile machine = new AssemblyFile(swApp);
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

    public static void Run2(string outFolder, SldWorks swApp)
    {
        if (string.IsNullOrWhiteSpace(outFolder))
            throw new ArgumentException("Output folder is required.", nameof(outFolder));

        SkeletonPart skeleton = new SkeletonPart(swApp);
        skeleton.OutputFolder = outFolder;
        skeleton.DESideOffset = 2.0;
        skeleton.NDESideOffset = 2.0;
        skeleton.GroundOffsetMm = -0.5;
        skeleton.CloseAfterCreate = true;

        StatorSheetPart statorSheet = new StatorSheetPart(swApp);
        statorSheet.OutputFolder = outFolder;
        statorSheet.CloseAfterCreate = true;

        ShaftPart shaft = new ShaftPart(swApp);
        shaft.OutputFolder = outFolder;
        shaft.CloseAfterCreate = true;

        AssemblyFile machine = new AssemblyFile(swApp);
        machine.FileName = "MachineAssembly.SLDASM";
        machine.OutputFolder = outFolder;

        string skeletonPath = skeleton.Create();
        string statorSheetPath = statorSheet.Create();
        string shaftPath = shaft.Create();
        string machinePath = machine.Create();

        var insertedSkeleton = machine.Insert(skeletonPath);
        machine.MateToOrigin(insertedSkeleton);
        var insertedStatorSheet = machine.Insert(statorSheetPath);
        var insertedShaft = machine.Insert(shaftPath);
        machine.MateCoincident(insertedSkeleton, "X-Achse", insertedStatorSheet, "Z-Achse");
        machine.MateCoincident(insertedSkeleton, "Ebene rechts", insertedStatorSheet, "Ebene vorne");
        machine.MateCoincident(insertedSkeleton, "Ebene vorne", insertedStatorSheet, "Ebene rechts");
        machine.MateCoincident(insertedSkeleton, "X-Achse", insertedShaft, "Z-Achse");

        Console.WriteLine($"Macro completed. Machine assembly: {machinePath}");
    }

    public static void Run3(string outFolder, SldWorks swApp)
    {
        if (string.IsNullOrWhiteSpace(outFolder))
            throw new ArgumentException("Output folder is required.", nameof(outFolder));

        _ = swApp;
        Console.WriteLine("Run3 is reserved for ad-hoc generator checks.");
    }

    public static void Run4(string outFolder, SldWorks swApp)
    {
        if (string.IsNullOrWhiteSpace(outFolder))
            throw new ArgumentException("Output folder is required.", nameof(outFolder));

        // These objects are the editable parameter surface for the machine build.
        // Skeleton planes and axes act as the assembly reference frame.
        SkeletonPart skeleton = new SkeletonPart(swApp);
        skeleton.OutputFolder = outFolder;
        skeleton.CloseAfterCreate = true;
        skeleton.LocalFileName = "skeleton.SLDPRT";
        skeleton.DESideOffset = 2.0;
        skeleton.NDESideOffset = 2.0;
        skeleton.GroundOffsetMm = -0.5;

        // Main stator lamination used as the thick core pack.
        StatorSheetPart statorSheet = new StatorSheetPart(swApp);
        statorSheet.OutputFolder = outFolder;
        statorSheet.CloseAfterCreate = true;
        statorSheet.LocalFileName = "StatorBleche.SLDPRT";
        statorSheet.OuterDiameterMm = 0.99;
        statorSheet.InnerDiameterMm = 0.64;
        statorSheet.PlateThicknessMm = 0.1;
        statorSheet.SlotWidthMm = 0.0157;
        statorSheet.SlotBottomYmm = 0.32;
        statorSheet.SlotTopYmm = 0.4052;
        statorSheet.AngleInDegrees = 70.0;
        statorSheet.FilletRadiusMm = 0.001;
        statorSheet.SlotGuideSpacingMm = 0.0057;
        statorSheet.SlotGuideOffsetMm = -0.0007;
        statorSheet.SlotPatternCount = 60;
        statorSheet.MaterialName = "AISI 1020";

        // Distance sheets add the repeated spacer/boss geometry between stator packs.
        StatorDistanceSheetPart statorDistanceSheet = new StatorDistanceSheetPart(swApp);
        statorDistanceSheet.OutputFolder = outFolder;
        statorDistanceSheet.CloseAfterCreate = true;
        statorDistanceSheet.LocalFileName = "StatorDistanceBleche.SLDPRT";
        statorDistanceSheet.OuterDiameterMm = 0.99;
        statorDistanceSheet.InnerDiameterMm = 0.64;
        statorDistanceSheet.PlateThicknessMm = 0.001;
        statorDistanceSheet.SlotWidthMm = 0.0205;
        statorDistanceSheet.SlotBottomYmm = 0.32;
        statorDistanceSheet.SlotTopYmm = 0.406;
        statorDistanceSheet.BossRectangleHeightMm = 0.16;
        statorDistanceSheet.BossRectangleWidthMm = 0.008;
        statorDistanceSheet.BossOuterDiameterOffsetMm = 0.009;
        statorDistanceSheet.BossCenterlineAngleDeg = 2.96;
        statorDistanceSheet.BossExtrusionDepthMm = 0.01;
        statorDistanceSheet.BossCutOuterTabWidthMm = 0.002;
        statorDistanceSheet.BossCutOuterTabHeightMm = 0.0025;
        statorDistanceSheet.BossCutTopShelfThicknessMm = 0.0015;
        statorDistanceSheet.BossCutInnerLegWidthMm = 0.0015;
        statorDistanceSheet.BossCutBoundaryExtensionMm = 0.002;
        statorDistanceSheet.BossCircularPatternCount = 60;
        statorDistanceSheet.SlotPatternCount = 60;
        statorDistanceSheet.MaterialName = "AISI 1020";

        // End sheets cap the repeated lamination stack at each side.
        StatorEndSheetPart statorEndSheet = new StatorEndSheetPart(swApp);
        statorEndSheet.OutputFolder = outFolder;
        statorEndSheet.CloseAfterCreate = true;
        statorEndSheet.LocalFileName = "StatorEndBleche.SLDPRT";
        statorEndSheet.OuterDiameterMm = 0.99;
        statorEndSheet.InnerDiameterMm = 0.64;
        statorEndSheet.PlateThicknessMm = 0.001;
        statorEndSheet.SlotWidthMm = 0.0205;
        statorEndSheet.SlotBottomYmm = 0.32;
        statorEndSheet.SlotTopYmm = 0.406;
        statorEndSheet.SlotPatternCount = 60;
        statorEndSheet.MaterialName = "AISI 1020";

        // Torsion bars are created once and then patterned around the finished stack.
        TorsionBarPart torsionBar = new TorsionBarPart(swApp);
        torsionBar.OutputFolder = outFolder;
        torsionBar.CloseAfterCreate = true;
        torsionBar.LocalFileName = "TorsionBar.SLDPRT";
        torsionBar.BarLengthMm = 1.074;
        torsionBar.BarHeightMm = 0.04;
        torsionBar.BarThicknessMm = 0.03;
        torsionBar.HoleCenterlineOffsetFromBottomMm = 0.02;
        torsionBar.OuterHoleEndOffsetMm = 0.03;
        torsionBar.HolePairSpacingMm = 0.315;
        torsionBar.OuterHoleDiameterMm = 0.01;
        torsionBar.InnerHoleDiameterMm = 0.016;
        torsionBar.CenterHoleDiameterMm = 0.016;
        torsionBar.OuterTapSizePrimary = "M10x1.5";
        torsionBar.OuterTapSizeFallback = "M10";
        torsionBar.InnerTapSizePrimary = "M16x2";
        torsionBar.InnerTapSizeFallback = "M16";
        torsionBar.P0001ConfigName = "P0001";
        torsionBar.P0002ConfigName = "P0002";
        torsionBar.MaterialName = "AISI 1020";

        // Press plates clamp the stack and also carry the assembly placement angle.
        PressPlatePart pressPlate = new PressPlatePart(swApp);
        pressPlate.OutputFolder = outFolder;
        pressPlate.CloseAfterCreate = true;
        pressPlate.LocalFileName = "PressPlate.SLDPRT";
        pressPlate.OuterDiameterMm = 0.99;
        pressPlate.RingInnerDiameterMm = 0.84;
        pressPlate.PlateOuterInsetFromOuterDiameterMm = 0.005;
        pressPlate.PlateRadialLengthMm = 0.165;
        pressPlate.RingThicknessMm = 0.002;
        pressPlate.PlateBodyThicknessMm = 0.01;
        pressPlate.PlateWidthMm = 0.006;
        pressPlate.PlateCount = 60;
        pressPlate.AssemblyAngleDeg = 3.0;
        pressPlate.MaterialName = "AISI 1020";

        // The NDE press ring is the first and last hardware item in the axial stack.
        StatorPressringNdePart pressRingNde = new StatorPressringNdePart(swApp);
        pressRingNde.OutputFolder = outFolder;
        pressRingNde.CloseAfterCreate = true;
        pressRingNde.LocalFileName = "StatorPressringNDE.SLDPRT";
        pressRingNde.OuterDiameterMm = 1.1;
        pressRingNde.InnerDiameterMm = 0.84;
        pressRingNde.PressRingOuterDiameterMm = 0.86;
        pressRingNde.RingThicknessMm = 0.028;
        pressRingNde.PressRingThicknessMm = 0.002;
        pressRingNde.BaseInnerChamferDistanceMm = 0.02;
        pressRingNde.BaseInnerChamferAngleDeg = 30.0;
        pressRingNde.PocketCenterRadiusMm = 0.52;
        pressRingNde.PocketWidthMm = 0.043;
        pressRingNde.PocketHeightMm = 0.036;
        pressRingNde.PocketCornerRadiusMm = 0.005;
        pressRingNde.PocketCount = 8;
        pressRingNde.MaterialName = "AISI 1020";

        // This is the target assembly document that all generated parts are inserted into.
        AssemblyFile machine = new AssemblyFile(swApp);
        machine.OutputFolder = outFolder;
        machine.FileName = "MachineAssembly.SLDASM";
        machine.CloseAfterCreate = false;
        int repeatedDistanceEndSheetPacks = 5;

        double AsMeters(double value) => value;

        // Keep stack spacing derived from the current part settings.
        double statorSheetPackThicknessMm = statorSheet.PlateThicknessMm;
        double statorDistanceSheetStackThicknessMm = statorDistanceSheet.PlateThicknessMm + statorDistanceSheet.BossExtrusionDepthMm;
        double statorEndSheetStackThicknessMm = statorEndSheet.PlateThicknessMm;
        double pressPlateStackThicknessMm = Math.Max(pressPlate.PlateBodyThicknessMm, pressPlate.RingThicknessMm);
        double pressRingNdeStackThicknessMm = pressRingNde.RingThicknessMm;
        double repeatedStackBlockThicknessMm = statorDistanceSheetStackThicknessMm + statorEndSheetStackThicknessMm + statorSheetPackThicknessMm;
        double torsionBarSlotRadiusMm = pressRingNde.PocketCenterRadiusMm;
        int torsionBarPatternCount = pressRingNde.PocketCount;

        // Create all source documents before starting assembly placement.
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

        double stackOffsetMm = 0.0;
        Component2 insertedStackComponent;

        insertedStackComponent = machine.Insert(pressRingNdePath);
        machine.MateCoincident(insertedSkeleton, "X-Achse", insertedStackComponent, "Z-Achse");
        machine.MateCoincident(insertedSkeleton, "Ebene vorne", insertedStackComponent, "Ebene rechts");
        machine.MateCoincident(insertedSkeleton, "Ebene rechts", insertedStackComponent, "Ebene vorne", AsMeters(stackOffsetMm));
        stackOffsetMm += pressRingNdeStackThicknessMm;

        insertedStackComponent = machine.Insert(pressPlatePath);
        machine.MateCoincident(insertedSkeleton, "X-Achse", insertedStackComponent, "Z-Achse");
        machine.MateAngle(insertedSkeleton, "Ebene vorne", insertedStackComponent, "Ebene rechts", pressPlate.AssemblyAngleDeg);
        machine.MateCoincident(insertedSkeleton, "Ebene rechts", insertedStackComponent, "Ebene vorne", AsMeters(stackOffsetMm));
        stackOffsetMm += pressPlateStackThicknessMm;

        insertedStackComponent = machine.Insert(statorEndSheetPath);
        machine.MateCoincident(insertedSkeleton, "X-Achse", insertedStackComponent, "Z-Achse");
        machine.MateCoincident(insertedSkeleton, "Ebene vorne", insertedStackComponent, "Ebene rechts");
        machine.MateCoincident(insertedSkeleton, "Ebene rechts", insertedStackComponent, "Ebene vorne", AsMeters(stackOffsetMm));
        stackOffsetMm += statorEndSheetStackThicknessMm;

        insertedStackComponent = machine.Insert(statorSheetPath);
        machine.MateCoincident(insertedSkeleton, "X-Achse", insertedStackComponent, "Z-Achse");
        machine.MateCoincident(insertedSkeleton, "Ebene vorne", insertedStackComponent, "Ebene rechts");
        machine.MateCoincident(insertedSkeleton, "Ebene rechts", insertedStackComponent, "Ebene vorne", AsMeters(stackOffsetMm));
        stackOffsetMm += statorSheetPackThicknessMm;

        Component2 repeatedDistanceSeed = machine.Insert(statorDistanceSheetPath);
        machine.MateCoincident(insertedSkeleton, "X-Achse", repeatedDistanceSeed, "Z-Achse");
        machine.MateCoincident(insertedSkeleton, "Ebene vorne", repeatedDistanceSeed, "Ebene rechts");
        machine.MateCoincident(insertedSkeleton, "Ebene rechts", repeatedDistanceSeed, "Ebene vorne", AsMeters(stackOffsetMm));
        stackOffsetMm += statorDistanceSheetStackThicknessMm;

        Component2 repeatedEndSeed = machine.Insert(statorEndSheetPath);
        machine.MateCoincident(insertedSkeleton, "X-Achse", repeatedEndSeed, "Z-Achse");
        machine.MateCoincident(insertedSkeleton, "Ebene vorne", repeatedEndSeed, "Ebene rechts");
        machine.MateCoincident(insertedSkeleton, "Ebene rechts", repeatedEndSeed, "Ebene vorne", AsMeters(stackOffsetMm));
        stackOffsetMm += statorEndSheetStackThicknessMm;

        Component2 repeatedStatorSeed = machine.Insert(statorSheetPath);
        machine.MateCoincident(insertedSkeleton, "X-Achse", repeatedStatorSeed, "Z-Achse");
        machine.MateCoincident(insertedSkeleton, "Ebene vorne", repeatedStatorSeed, "Ebene rechts");
        machine.MateCoincident(insertedSkeleton, "Ebene rechts", repeatedStatorSeed, "Ebene vorne", AsMeters(stackOffsetMm));
        stackOffsetMm += statorSheetPackThicknessMm;

        // Repeat the middle stator block along the machine length.
        machine.LinearPattern(
            insertedSkeleton,
            "X-Achse",
            repeatedDistanceEndSheetPacks + 1,
            AsMeters(repeatedStackBlockThicknessMm),
            repeatedDistanceSeed,
            repeatedEndSeed,
            repeatedStatorSeed);

        // Close the far end with the same hardware sequence used at the start.
        stackOffsetMm += repeatedDistanceEndSheetPacks * repeatedStackBlockThicknessMm;

        insertedStackComponent = machine.Insert(statorEndSheetPath);
        machine.MateCoincident(insertedSkeleton, "X-Achse", insertedStackComponent, "Z-Achse");
        machine.MateCoincident(insertedSkeleton, "Ebene vorne", insertedStackComponent, "Ebene rechts", 0, true);
        machine.MateCoincident(insertedSkeleton, "Ebene rechts", insertedStackComponent, "Ebene vorne", AsMeters(stackOffsetMm));
        stackOffsetMm += statorEndSheetStackThicknessMm;

        insertedStackComponent = machine.Insert(pressPlatePath);
        machine.MateCoincident(insertedSkeleton, "X-Achse", insertedStackComponent, "Z-Achse");
        machine.MateAngle(insertedSkeleton, "Ebene vorne", insertedStackComponent, "Ebene rechts", pressPlate.AssemblyAngleDeg);
        machine.MateCoincident(insertedSkeleton, "Ebene rechts", insertedStackComponent, "Ebene vorne", AsMeters(stackOffsetMm));
        stackOffsetMm += pressPlateStackThicknessMm;

        insertedStackComponent = machine.Insert(pressRingNdePath);
        machine.MateCoincident(insertedSkeleton, "X-Achse", insertedStackComponent, "Z-Achse");
        machine.MateCoincident(insertedSkeleton, "Ebene vorne", insertedStackComponent, "Ebene rechts");
        machine.MateCoincident(insertedSkeleton, "Ebene rechts", insertedStackComponent, "Ebene vorne", AsMeters(stackOffsetMm));
        stackOffsetMm += pressRingNdeStackThicknessMm;

        // Place one torsion bar and pattern it around the main axis.
        Component2 insertedTorsionBar = machine.Insert(torsionBarPath);
        machine.MateParallel(insertedSkeleton, "X-Achse", insertedTorsionBar, "X-Achse");
        machine.MateCoincident(insertedSkeleton, "Ebene vorne", insertedTorsionBar, "Ebene oben");
        machine.MateCoincident(insertedSkeleton, "Ebene oben", insertedTorsionBar, "Ebene vorne", AsMeters(torsionBarSlotRadiusMm));
        machine.MateCoincident(insertedSkeleton, "Ebene rechts", insertedTorsionBar, "Ebene rechts", AsMeters(stackOffsetMm / 2.0));
        machine.CircularPattern(insertedSkeleton, "X-Achse", torsionBarPatternCount, 2 * Math.PI, insertedTorsionBar);

        Console.WriteLine($"Macro completed. Machine assembly: {machinePath}");
    }

    public static void Run5(string outFolder, SldWorks swApp)
    {
        if (string.IsNullOrWhiteSpace(outFolder))
            throw new ArgumentException("Output folder is required.", nameof(outFolder));

        // This is the object you edit.
        // It owns everything for the torsion bar:
        // 1. the 3D model parameters
        // 2. the 2D drawing parameters
        TorsionBarPart torsionBar = new TorsionBarPart(swApp);
        torsionBar.OutputFolder = outFolder;
        torsionBar.CloseAfterCreate = true;
        torsionBar.LocalFileName = "TorsionBar.SLDPRT";
        torsionBar.BarLengthMm = 1.074;
        torsionBar.BarHeightMm = 0.04;
        torsionBar.BarThicknessMm = 0.03;
        torsionBar.HoleCenterlineOffsetFromBottomMm = 0.02;
        torsionBar.OuterHoleEndOffsetMm = 0.03;
        torsionBar.HolePairSpacingMm = 0.315;
        torsionBar.OuterHoleDiameterMm = 0.01;
        torsionBar.InnerHoleDiameterMm = 0.016;
        torsionBar.CenterHoleDiameterMm = 0.016;
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
        // false = leave the drawing open after Create(); true = close it after saving
        torsionBar.DrawingCloseAfterCreate = false;
        torsionBar.DrawingLocalFileName = "TorsionBar.SLDDRW";
        torsionBar.DrawingSheetName = "Torsion Bar";
        torsionBar.DrawingLanguageCode = "EN";
        torsionBar.DrawingTemplateFolderPath = string.Empty;
        torsionBar.DrawingBottomTitleBlockClearanceMm = 0.085;
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
