// Import base .NET types; alternative: remove if unused to speed build slightly.
using System;

namespace SwAutomation;

// User-friendly face names that map to rectangular-part face directions.
// Note: "Up" maps to "Top" and "Down" maps to "Bottom".
public enum FaceName
{
    Top,
    Bottom,
    Left,
    Right,
    Front,
    Back,
    Up,
    Down
}

// Sketch-plane options for part base sketch.
// "Side" is treated as alias of Right plane.
public enum SketchPlaneName
{
    Front,
    Top,
    Right,
    Side
}

// Holds one mate instruction: face on part A mated to face on part B.
public sealed class FaceMatePair
{
    // Face selector for component A (named face or XYZ point).
    public FaceSelection PartAFace { get; }

    // Face selector for component B (named face or XYZ point).
    public FaceSelection PartBFace { get; }

    // True = use flipped (anti-aligned) mate direction; false = normal aligned.
    public bool IsFlipped { get; }

    // Construct one face-to-face mate instruction using named faces.
    public FaceMatePair(FaceName partAFace, FaceName partBFace)
        : this(new FaceSelection(partAFace), new FaceSelection(partBFace), false)
    {
    }

    // Construct one face-to-face mate instruction using named faces with flip option.
    public FaceMatePair(FaceName partAFace, FaceName partBFace, bool isFlipped)
        : this(new FaceSelection(partAFace), new FaceSelection(partBFace), isFlipped)
    {
    }

    // Construct one face-to-face mate instruction using selectors.
    public FaceMatePair(FaceSelection partAFace, FaceSelection partBFace, bool isFlipped = false)
    {
        // Store selector for component A.
        PartAFace = partAFace;
        // Store selector for component B.
        PartBFace = partBFace;
        // Store mate direction preference.
        IsFlipped = isFlipped;
    }
}

// Describes how to pick one face: by FaceName or by XYZ point (from part origin, in mm).
public readonly struct FaceSelection
{
    // True when selection should use XYZ point mode.
    public bool IsPointSelection { get; }

    // Named-face mode value (used when IsPointSelection=false).
    public FaceName FaceName { get; }

    // Point mode X (mm from part origin).
    public double XMm { get; }

    // Point mode Y (mm from part origin).
    public double YMm { get; }

    // Point mode Z (mm from part origin).
    public double ZMm { get; }

    // Build a named-face selector.
    public FaceSelection(FaceName faceName)
    {
        IsPointSelection = false;
        FaceName = faceName;
        XMm = 0;
        YMm = 0;
        ZMm = 0;
    }

    // Build a point selector (coordinates are from part origin in mm).
    public FaceSelection(double xMm, double yMm, double zMm)
    {
        IsPointSelection = true;
        FaceName = FaceName.Top;
        XMm = xMm;
        YMm = yMm;
        ZMm = zMm;
    }

    // Readable text for logging.
    public string ToDisplayText()
    {
        if (IsPointSelection) return "Point(" + XMm + ", " + YMm + ", " + ZMm + ") mm";
        return FaceName.ToString();
    }
}

// Collects part-generation inputs so long argument lists are replaced by one object.
public readonly struct PartParameters
{
    // Part file name without extension.
    public string Name { get; }

    // Rectangle width in mm.
    public double WidthMm { get; }

    // Rectangle depth in mm.
    public double DepthMm { get; }

    // Extrusion height in mm.
    public double HeightMm { get; }

    // Optional hole diameter in mm.
    public double HoleDiameterMm { get; }

    // Optional linear hole count.
    public int HoleCount { get; }

    // Base sketch plane selection (Front is default).
    public SketchPlaneName SketchPlane { get; }

    // Build one immutable parameter object.
    public PartParameters(string name, double widthMm, double depthMm, double heightMm, double holeDiameterMm, int holeCount, SketchPlaneName sketchPlane = SketchPlaneName.Front)
    {
        // Store all values as-is; validation is handled by PartBuilder.
        Name = name;
        WidthMm = widthMm;
        DepthMm = depthMm;
        HeightMm = heightMm;
        HoleDiameterMm = holeDiameterMm;
        HoleCount = holeCount;
        SketchPlane = sketchPlane;
    }
}
