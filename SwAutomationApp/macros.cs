using System;
using SolidWorks.Interop.sldworks;
using SwAutomation.Pdm;

namespace SwAutomation;

public static class Project1
{
    public static void Run(string outFolder, SldWorks swApp, PdmModule pdm)
    {
        if (string.IsNullOrWhiteSpace(outFolder))
            throw new ArgumentException("Output folder is required.", nameof(outFolder));

        SkeletonPart skeleton = new SkeletonPart(swApp, pdm);
        skeleton.OutputFolder = outFolder;
        skeleton.SideOffsetMm = 500.0;
        skeleton.GroundOffsetMm = -250.0;
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

    public static void Run2(string outFolder, SldWorks swApp, PdmModule pdm)
    {
        if (string.IsNullOrWhiteSpace(outFolder))
            throw new ArgumentException("Output folder is required.", nameof(outFolder));

        SkeletonPart skeleton = new SkeletonPart(swApp, pdm);
        skeleton.OutputFolder = outFolder;
        skeleton.SideOffsetMm = 2000.0;
        skeleton.GroundOffsetMm = -500.0;
        skeleton.CloseAfterCreate = true;

        StatorSheetPart statorSheet = new StatorSheetPart(swApp, pdm);
        statorSheet.OutputFolder = outFolder;
        statorSheet.CloseAfterCreate = true;

        ShaftPart shaft = new ShaftPart(swApp, pdm);
        shaft.OutputFolder = outFolder;
        shaft.CloseAfterCreate = true;

        AssemblyFile machine = new AssemblyFile(swApp, pdm);
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

    public static void Run3(string outFolder, SldWorks swApp, PdmModule pdm)
    {
        if (string.IsNullOrWhiteSpace(outFolder))
            throw new ArgumentException("Output folder is required.", nameof(outFolder));

        _ = swApp;
        _ = pdm;
        Console.WriteLine("Run3 is reserved for ad-hoc generator checks.");
    }

    public static void Run4(string outFolder, SldWorks swApp, PdmModule pdm)
    {
        if (string.IsNullOrWhiteSpace(outFolder))
            throw new ArgumentException("Output folder is required.", nameof(outFolder));

        // These objects are the editable parameter surface for the machine build.
        // Skeleton planes and axes act as the assembly reference frame.
        SkeletonPart skeleton = new SkeletonPart(swApp, pdm);
        skeleton.OutputFolder = outFolder;
        skeleton.SaveToPdm = false;
        skeleton.CloseAfterCreate = true;
        skeleton.LocalFileName = "skeleton.SLDPRT";
        skeleton.SideOffsetMm = 2000.0;
        skeleton.GroundOffsetMm = -500.0;
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

        // Main stator lamination used as the thick core pack.
        StatorSheetPart statorSheet = new StatorSheetPart(swApp, pdm);
        statorSheet.OutputFolder = outFolder;
        statorSheet.SaveToPdm = false;
        statorSheet.CloseAfterCreate = true;
        statorSheet.LocalFileName = "StatorBleche.SLDPRT";
        statorSheet.OuterDiameterMm = 990.0;
        statorSheet.InnerDiameterMm = 640.0;
        statorSheet.PlateThicknessMm = 100.0;
        statorSheet.SlotWidthMm = 15.7;
        statorSheet.SlotBottomYmm = 320.0;
        statorSheet.SlotTopYmm = 405.2;
        statorSheet.AngleInDegrees = 70.0;
        statorSheet.FilletRadiusMm = 1.0;
        statorSheet.SlotGuideSpacingMm = 5.7;
        statorSheet.SlotGuideOffsetMm = -0.7;
        statorSheet.SlotPatternCount = 60;
        statorSheet.MaterialName = "AISI 1020";

        // Distance sheets add the repeated spacer/boss geometry between stator packs.
        StatorDistanceSheetPart statorDistanceSheet = new StatorDistanceSheetPart(swApp, pdm);
        statorDistanceSheet.OutputFolder = outFolder;
        statorDistanceSheet.SaveToPdm = false;
        statorDistanceSheet.CloseAfterCreate = true;
        statorDistanceSheet.LocalFileName = "StatorDistanceBleche.SLDPRT";
        statorDistanceSheet.OuterDiameterMm = 990.0;
        statorDistanceSheet.InnerDiameterMm = 640.0;
        statorDistanceSheet.PlateThicknessMm = 1.0;
        statorDistanceSheet.SlotWidthMm = 20.5;
        statorDistanceSheet.SlotBottomYmm = 320.0;
        statorDistanceSheet.SlotTopYmm = 406.0;
        statorDistanceSheet.BossRectangleHeightMm = 160.0;
        statorDistanceSheet.BossRectangleWidthMm = 8.0;
        statorDistanceSheet.BossOuterDiameterOffsetMm = 9.0;
        statorDistanceSheet.BossCenterlineAngleDeg = 2.96;
        statorDistanceSheet.BossExtrusionDepthMm = 10.0;
        statorDistanceSheet.BossCutOuterTabWidthMm = 2.0;
        statorDistanceSheet.BossCutOuterTabHeightMm = 2.5;
        statorDistanceSheet.BossCutTopShelfThicknessMm = 1.5;
        statorDistanceSheet.BossCutInnerLegWidthMm = 1.5;
        statorDistanceSheet.BossCutBoundaryExtensionMm = 2.0;
        statorDistanceSheet.BossCircularPatternCount = 60;
        statorDistanceSheet.SlotPatternCount = 60;
        statorDistanceSheet.MaterialName = "AISI 1020";

        // End sheets cap the repeated lamination stack at each side.
        StatorEndSheetPart statorEndSheet = new StatorEndSheetPart(swApp, pdm);
        statorEndSheet.OutputFolder = outFolder;
        statorEndSheet.SaveToPdm = false;
        statorEndSheet.CloseAfterCreate = true;
        statorEndSheet.LocalFileName = "StatorEndBleche.SLDPRT";
        statorEndSheet.OuterDiameterMm = 990.0;
        statorEndSheet.InnerDiameterMm = 640.0;
        statorEndSheet.PlateThicknessMm = 1.0;
        statorEndSheet.SlotWidthMm = 20.5;
        statorEndSheet.SlotBottomYmm = 320.0;
        statorEndSheet.SlotTopYmm = 406.0;
        statorEndSheet.SlotPatternCount = 60;
        statorEndSheet.MaterialName = "AISI 1020";

        // Torsion bars are created once and then patterned around the finished stack.
        TorsionBarPart torsionBar = new TorsionBarPart(swApp, pdm);
        torsionBar.OutputFolder = outFolder;
        torsionBar.SaveToPdm = false;
        torsionBar.CloseAfterCreate = true;
        torsionBar.LocalFileName = "TorsionBar.SLDPRT";
        torsionBar.BarLengthMm = 1074.0;
        torsionBar.BarHeightMm = 40.0;
        torsionBar.BarThicknessMm = 30.0;
        torsionBar.HoleCenterlineOffsetFromBottomMm = 20.0;
        torsionBar.OuterHoleEndOffsetMm = 30.0;
        torsionBar.HolePairSpacingMm = 315.0;
        torsionBar.OuterHoleDiameterMm = 10.0;
        torsionBar.InnerHoleDiameterMm = 16.0;
        torsionBar.CenterHoleDiameterMm = 16.0;
        torsionBar.OuterTapSizePrimary = "M10x1.5";
        torsionBar.OuterTapSizeFallback = "M10";
        torsionBar.InnerTapSizePrimary = "M16x2";
        torsionBar.InnerTapSizeFallback = "M16";
        torsionBar.P0001ConfigName = "P0001";
        torsionBar.P0002ConfigName = "P0002";
        torsionBar.MaterialName = "AISI 1020";

        // Press plates clamp the stack and also carry the assembly placement angle.
        PressPlatePart pressPlate = new PressPlatePart(swApp, pdm);
        pressPlate.OutputFolder = outFolder;
        pressPlate.SaveToPdm = false;
        pressPlate.CloseAfterCreate = true;
        pressPlate.LocalFileName = "PressPlate.SLDPRT";
        pressPlate.OuterDiameterMm = 990.0;
        pressPlate.RingInnerDiameterMm = 840.0;
        pressPlate.PlateOuterInsetFromOuterDiameterMm = 5.0;
        pressPlate.PlateRadialLengthMm = 165.0;
        pressPlate.RingThicknessMm = 2.0;
        pressPlate.PlateBodyThicknessMm = 10.0;
        pressPlate.PlateWidthMm = 6.0;
        pressPlate.PlateCount = 60;
        pressPlate.AssemblyAngleDeg = 3.0;
        pressPlate.MaterialName = "AISI 1020";

        // The NDE press ring is the first and last hardware item in the axial stack.
        StatorPressringNdePart pressRingNde = new StatorPressringNdePart(swApp, pdm);
        pressRingNde.OutputFolder = outFolder;
        pressRingNde.SaveToPdm = false;
        pressRingNde.CloseAfterCreate = true;
        pressRingNde.LocalFileName = "StatorPressringNDE.SLDPRT";
        pressRingNde.OuterDiameterMm = 1100.0;
        pressRingNde.InnerDiameterMm = 840.0;
        pressRingNde.PressRingOuterDiameterMm = 860.0;
        pressRingNde.RingThicknessMm = 28.0;
        pressRingNde.PressRingThicknessMm = 2.0;
        pressRingNde.BaseInnerChamferDistanceMm = 20.0;
        pressRingNde.BaseInnerChamferAngleDeg = 30.0;
        pressRingNde.PocketCenterRadiusMm = 520.0;
        pressRingNde.PocketWidthMm = 43.0;
        pressRingNde.PocketHeightMm = 36.0;
        pressRingNde.PocketCornerRadiusMm = 5.0;
        pressRingNde.PocketCount = 8;
        pressRingNde.MaterialName = "AISI 1020";

        // This is the target assembly document that all generated parts are inserted into.
        AssemblyFile machine = new AssemblyFile(swApp, pdm);
        machine.OutputFolder = outFolder;
        machine.SaveToPdm = false;
        machine.FileName = "MachineAssembly.SLDASM";
        machine.CloseAfterCreate = false;
        int repeatedDistanceEndSheetPacks = 5;

        double Mm(double mm) => mm / 1000.0;

        // These values are derived from the editable object settings above.
        // If you change part thicknesses or pocket counts, the assembly spacing below follows automatically.
        double statorSheetPackThicknessMm = statorSheet.PlateThicknessMm;
        double statorDistanceSheetStackThicknessMm = statorDistanceSheet.PlateThicknessMm + statorDistanceSheet.BossExtrusionDepthMm;
        double statorEndSheetStackThicknessMm = statorEndSheet.PlateThicknessMm;
        double pressPlateStackThicknessMm = Math.Max(pressPlate.PlateBodyThicknessMm, pressPlate.RingThicknessMm);
        double pressRingNdeStackThicknessMm = pressRingNde.RingThicknessMm;
        double repeatedStackBlockThicknessMm = statorDistanceSheetStackThicknessMm + statorEndSheetStackThicknessMm + statorSheetPackThicknessMm;
        double torsionBarSlotRadiusMm = pressRingNde.PocketCenterRadiusMm;
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

        double stackOffsetMm = 0.0;
        Component2 insertedStackComponent;

        insertedStackComponent = machine.Insert(pressRingNdePath);
        machine.MateCoincident(insertedSkeleton, "X-Achse", insertedStackComponent, "Z-Achse");
        machine.MateCoincident(insertedSkeleton, "Ebene vorne", insertedStackComponent, "Ebene rechts");
        machine.MateCoincident(insertedSkeleton, "Ebene rechts", insertedStackComponent, "Ebene vorne", Mm(stackOffsetMm));
        stackOffsetMm += pressRingNdeStackThicknessMm;

        insertedStackComponent = machine.Insert(pressPlatePath);
        machine.MateCoincident(insertedSkeleton, "X-Achse", insertedStackComponent, "Z-Achse");
        machine.MateAngle(insertedSkeleton, "Ebene vorne", insertedStackComponent, "Ebene rechts", pressPlate.AssemblyAngleDeg);
        machine.MateCoincident(insertedSkeleton, "Ebene rechts", insertedStackComponent, "Ebene vorne", Mm(stackOffsetMm));
        stackOffsetMm += pressPlateStackThicknessMm;

        insertedStackComponent = machine.Insert(statorEndSheetPath);
        machine.MateCoincident(insertedSkeleton, "X-Achse", insertedStackComponent, "Z-Achse");
        machine.MateCoincident(insertedSkeleton, "Ebene vorne", insertedStackComponent, "Ebene rechts");
        machine.MateCoincident(insertedSkeleton, "Ebene rechts", insertedStackComponent, "Ebene vorne", Mm(stackOffsetMm));
        stackOffsetMm += statorEndSheetStackThicknessMm;

        insertedStackComponent = machine.Insert(statorSheetPath);
        machine.MateCoincident(insertedSkeleton, "X-Achse", insertedStackComponent, "Z-Achse");
        machine.MateCoincident(insertedSkeleton, "Ebene vorne", insertedStackComponent, "Ebene rechts");
        machine.MateCoincident(insertedSkeleton, "Ebene rechts", insertedStackComponent, "Ebene vorne", Mm(stackOffsetMm));
        stackOffsetMm += statorSheetPackThicknessMm;

        Component2 repeatedDistanceSeed = machine.Insert(statorDistanceSheetPath);
        machine.MateCoincident(insertedSkeleton, "X-Achse", repeatedDistanceSeed, "Z-Achse");
        machine.MateCoincident(insertedSkeleton, "Ebene vorne", repeatedDistanceSeed, "Ebene rechts");
        machine.MateCoincident(insertedSkeleton, "Ebene rechts", repeatedDistanceSeed, "Ebene vorne", Mm(stackOffsetMm));
        stackOffsetMm += statorDistanceSheetStackThicknessMm;

        Component2 repeatedEndSeed = machine.Insert(statorEndSheetPath);
        machine.MateCoincident(insertedSkeleton, "X-Achse", repeatedEndSeed, "Z-Achse");
        machine.MateCoincident(insertedSkeleton, "Ebene vorne", repeatedEndSeed, "Ebene rechts");
        machine.MateCoincident(insertedSkeleton, "Ebene rechts", repeatedEndSeed, "Ebene vorne", Mm(stackOffsetMm));
        stackOffsetMm += statorEndSheetStackThicknessMm;

        Component2 repeatedStatorSeed = machine.Insert(statorSheetPath);
        machine.MateCoincident(insertedSkeleton, "X-Achse", repeatedStatorSeed, "Z-Achse");
        machine.MateCoincident(insertedSkeleton, "Ebene vorne", repeatedStatorSeed, "Ebene rechts");
        machine.MateCoincident(insertedSkeleton, "Ebene rechts", repeatedStatorSeed, "Ebene vorne", Mm(stackOffsetMm));
        stackOffsetMm += statorSheetPackThicknessMm;

        // Repeat the middle stator block along the machine length.
        machine.LinearPattern(
            insertedSkeleton,
            "X-Achse",
            repeatedDistanceEndSheetPacks + 1,
            Mm(repeatedStackBlockThicknessMm),
            repeatedDistanceSeed,
            repeatedEndSeed,
            repeatedStatorSeed);

        // Close the far end with the same hardware sequence used at the start.
        stackOffsetMm += repeatedDistanceEndSheetPacks * repeatedStackBlockThicknessMm;

        insertedStackComponent = machine.Insert(statorEndSheetPath);
        machine.MateCoincident(insertedSkeleton, "X-Achse", insertedStackComponent, "Z-Achse");
        machine.MateCoincident(insertedSkeleton, "Ebene vorne", insertedStackComponent, "Ebene rechts", 0, true);
        machine.MateCoincident(insertedSkeleton, "Ebene rechts", insertedStackComponent, "Ebene vorne", Mm(stackOffsetMm));
        stackOffsetMm += statorEndSheetStackThicknessMm;

        insertedStackComponent = machine.Insert(pressPlatePath);
        machine.MateCoincident(insertedSkeleton, "X-Achse", insertedStackComponent, "Z-Achse");
        machine.MateAngle(insertedSkeleton, "Ebene vorne", insertedStackComponent, "Ebene rechts", pressPlate.AssemblyAngleDeg);
        machine.MateCoincident(insertedSkeleton, "Ebene rechts", insertedStackComponent, "Ebene vorne", Mm(stackOffsetMm));
        stackOffsetMm += pressPlateStackThicknessMm;

        insertedStackComponent = machine.Insert(pressRingNdePath);
        machine.MateCoincident(insertedSkeleton, "X-Achse", insertedStackComponent, "Z-Achse");
        machine.MateCoincident(insertedSkeleton, "Ebene vorne", insertedStackComponent, "Ebene rechts");
        machine.MateCoincident(insertedSkeleton, "Ebene rechts", insertedStackComponent, "Ebene vorne", Mm(stackOffsetMm));
        stackOffsetMm += pressRingNdeStackThicknessMm;

        // Place one torsion bar and pattern it around the main axis.
        Component2 insertedTorsionBar = machine.Insert(torsionBarPath);
        machine.MateParallel(insertedSkeleton, "X-Achse", insertedTorsionBar, "X-Achse");
        machine.MateCoincident(insertedSkeleton, "Ebene vorne", insertedTorsionBar, "Ebene oben");
        machine.MateCoincident(insertedSkeleton, "Ebene oben", insertedTorsionBar, "Ebene vorne", Mm(torsionBarSlotRadiusMm));
        machine.MateCoincident(insertedSkeleton, "Ebene rechts", insertedTorsionBar, "Ebene rechts", Mm(stackOffsetMm / 2.0));
        machine.CircularPattern(insertedSkeleton, "X-Achse", torsionBarPatternCount, 2 * Math.PI, insertedTorsionBar);

        Console.WriteLine($"Macro completed. Machine assembly: {machinePath}");
    }

    public static void Run5(string outFolder, SldWorks swApp, PdmModule pdm)
    {
        if (string.IsNullOrWhiteSpace(outFolder))
            throw new ArgumentException("Output folder is required.", nameof(outFolder));

        // This is the object you edit for the drawing test.
        // It owns everything for the torsion bar:
        // 1. the 3D model parameters
        // 2. the 2D drawing parameters
        TorsionBarPart torsionBar = new TorsionBarPart(swApp, pdm);
        torsionBar.OutputFolder = outFolder;
        torsionBar.SaveToPdm = false;
        torsionBar.CloseAfterCreate = true;
        torsionBar.LocalFileName = "TorsionBar.SLDPRT";
        torsionBar.BarLengthMm = 1074.0;
        torsionBar.BarHeightMm = 40.0;
        torsionBar.BarThicknessMm = 30.0;
        torsionBar.HoleCenterlineOffsetFromBottomMm = 20.0;
        torsionBar.OuterHoleEndOffsetMm = 30.0;
        torsionBar.HolePairSpacingMm = 315.0;
        torsionBar.OuterHoleDiameterMm = 10.0;
        torsionBar.InnerHoleDiameterMm = 16.0;
        torsionBar.CenterHoleDiameterMm = 16.0;
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
        torsionBar.DrawingSaveToPdm = false;
        // false = leave the drawing open after Create(); true = close it after saving
        torsionBar.DrawingCloseAfterCreate = false;
        torsionBar.DrawingLocalFileName = "TorsionBar.SLDDRW";
        torsionBar.DrawingSheetName = "Torsion Bar";
        torsionBar.DrawingLanguageCode = "EN";
        torsionBar.DrawingTemplateFolderPath = @"C:\Users\kareem.salah\PDM\Birr Machines PDM\40_Templates\Solidworks\Blattformate\Birr Machines";
        torsionBar.DrawingBottomTitleBlockClearanceMm = 85.0;
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
