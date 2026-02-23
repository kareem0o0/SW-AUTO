// Import base .NET types; alternative: remove if unused to speed build slightly.
using System;
// Import file system helpers for folders and paths.
using System.IO;
// Import SOLIDWORKS main COM interop types.
using SolidWorks.Interop.sldworks;
// Import SOLIDWORKS constant enums used for API calls.
using SolidWorks.Interop.swconst;

namespace SwAutomation;

// Contains assembly insertion, ray-cast face selection, interference checks, and mating.
public sealed class AssemblyBuilder
{
    // SOLIDWORKS app instance used by all assembly operations.
    private readonly SldWorks _swApp;

    // Inject SOLIDWORKS session dependency.
    public AssemblyBuilder(SldWorks swApp)
    {
        _swApp = swApp ?? throw new ArgumentNullException(nameof(swApp));
    }

    // Build an assembly that inserts two parts and mates planes.
    public void GenerateAssembly(string partAPath, string partBPath, string folder)
    {
        try
        {
            bool oldInputDimVal = _swApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate);
            _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

            string template = _swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateAssembly);
            ModelDoc2 assyModel = (ModelDoc2)_swApp.NewDocument(template, 0, 0, 0);
            AssemblyDoc swAssy = (AssemblyDoc)assyModel;

            Component2 compA = swAssy.AddComponent5(partAPath, (int)swAddComponentConfigOptions_e.swAddComponentConfigOptions_CurrentSelectedConfig, "", false, "", 0, 0, 0);
            Component2 compB = swAssy.AddComponent5(partBPath, (int)swAddComponentConfigOptions_e.swAddComponentConfigOptions_CurrentSelectedConfig, "", false, "", 0, 0, 0.05);

            string assemblyName = assyModel.GetTitle();
            if (assemblyName.EndsWith(".sldasm", StringComparison.OrdinalIgnoreCase))
            {
                assemblyName = assemblyName.Substring(0, assemblyName.Length - 7);
            }

            Console.WriteLine("Assembly component A instance: " + compA.Name2);
            Console.WriteLine("Assembly component B instance: " + compB.Name2);
            LogComponentReferencePlanes(compA, "A");
            LogComponentReferencePlanes(compB, "B");

            bool statusA;
            bool statusB;
            assyModel.ClearSelection2(true);
            bool frontPairSelected = SelectAssemblyPlanePair(assyModel, compA, compB, assemblyName, new[] { "Front Plane", "Front", "Ebene Vorne", "Vorne" }, 0, out statusA, out statusB);
            Console.WriteLine("Front plane select status A=" + statusA + ", B=" + statusB);
            if (frontPairSelected)
            {
                int mateError = 0;
                swAssy.AddMate3((int)swMateType_e.swMateCOINCIDENT, (int)swMateAlign_e.swMateAlignALIGNED, false, 0, 0, 0, 0, 0, 0, 0, 0, false, out mateError);
            }

            assyModel.ClearSelection2(true);
            bool topPairSelected = SelectAssemblyPlanePair(assyModel, compA, compB, assemblyName, new[] { "Top Plane", "Top", "Ebene Oben", "Oben" }, 1, out statusA, out statusB);
            Console.WriteLine("Top plane select status A=" + statusA + ", B=" + statusB);
            if (topPairSelected)
            {
                int mateError = 0;
                swAssy.AddMate3((int)swMateType_e.swMateCOINCIDENT, (int)swMateAlign_e.swMateAlignALIGNED, false, 0, 0, 0, 0, 0, 0, 0, 0, false, out mateError);
            }

            _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, oldInputDimVal);
            assyModel.ForceRebuild3(false);
            string assemblyPath = Path.Combine(folder, "Final_Assembly.SLDASM");
            assyModel.SaveAs3(assemblyPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Assembly generation by planes failed.", ex);
        }
    }

    // Build an assembly that inserts two parts and mates by physical faces.
    public void GenerateAssemblyByFaces(string partAPath, string partBPath, string folder)
    {
        FaceMatePair[] defaultPairs =
        {
            new FaceMatePair(FaceName.Top, FaceName.Bottom),
            new FaceMatePair(FaceName.Right, FaceName.Right),
            new FaceMatePair(FaceName.Front, FaceName.Front)
        };

        GenerateAssemblyByCustomFacePairs(partAPath, partBPath, folder, defaultPairs, "Final_Assembly_FaceMates.SLDASM");
    }


    // Build an assembly where user chooses exactly which face on A mates to which face on B.
    public void GenerateAssemblyByCustomFacePairs(string partAPath, string partBPath, string folder, FaceMatePair[] matePairs, string outputAssemblyFileName)
    {
        try
        {
            if (matePairs == null || matePairs.Length == 0)
            {
                Console.WriteLine("Custom face-mate assembly skipped: no mate pairs were provided.");
                return;
            }

            bool oldInputDimVal = _swApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate);
            _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

            string template = _swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateAssembly);
            ModelDoc2 assyModel = (ModelDoc2)_swApp.NewDocument(template, 0, 0, 0);
            AssemblyDoc swAssy = (AssemblyDoc)assyModel;

            Component2 compA = swAssy.AddComponent5(partAPath, (int)swAddComponentConfigOptions_e.swAddComponentConfigOptions_CurrentSelectedConfig, "", false, "", 0, 0, 0);
            Component2 compB = swAssy.AddComponent5(partBPath, (int)swAddComponentConfigOptions_e.swAddComponentConfigOptions_CurrentSelectedConfig, "", false, "", 0, 0, 0.08);

            if (compA == null || compB == null)
            {
                Console.WriteLine("Face-mate assembly failed: one or both components could not be inserted.");
                _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, oldInputDimVal);
                return;
            }

            Console.WriteLine("Face-mate assembly component A: " + compA.Name2);
            Console.WriteLine("Face-mate assembly component B: " + compB.Name2);

            int successCount = 0;
            int failCount = 0;

            for (int i = 0; i < matePairs.Length; i++)
            {
                FaceMatePair pair = matePairs[i];
                if (pair == null)
                {
                    Console.WriteLine("Face mate pair #" + (i + 1) + " skipped: pair is null.");
                    failCount++;
                    continue;
                }

                // Resolve selector A to a concrete face.
                Face2 targetFaceA = null;
                string faceALabel = pair.PartAFace.ToDisplayText();
                if (pair.PartAFace.IsPointSelection)
                {
                    targetFaceA = SelectComponentFaceByPartPoint(assyModel, compA, pair.PartAFace.XMm, pair.PartAFace.YMm, pair.PartAFace.ZMm);
                }
                else
                {
                    string faceAKey = NormalizeFaceName(pair.PartAFace.FaceName.ToString());
                    if (string.IsNullOrWhiteSpace(faceAKey))
                    {
                        Console.WriteLine("Face mate pair #" + (i + 1) + " failed: invalid face selector for part A.");
                        Console.WriteLine("Allowed face names: " + string.Join(", ", Enum.GetNames(typeof(FaceName))));
                        failCount++;
                        continue;
                    }

                    targetFaceA = SelectComponentFaceByRay(assyModel, compA, faceAKey);
                }

                // Resolve selector B to a concrete face.
                Face2 targetFaceB = null;
                string faceBLabel = pair.PartBFace.ToDisplayText();
                if (pair.PartBFace.IsPointSelection)
                {
                    targetFaceB = SelectComponentFaceByPartPoint(assyModel, compB, pair.PartBFace.XMm, pair.PartBFace.YMm, pair.PartBFace.ZMm);
                }
                else
                {
                    string faceBKey = NormalizeFaceName(pair.PartBFace.FaceName.ToString());
                    if (string.IsNullOrWhiteSpace(faceBKey))
                    {
                        Console.WriteLine("Face mate pair #" + (i + 1) + " failed: invalid face selector for part B.");
                        Console.WriteLine("Allowed face names: " + string.Join(", ", Enum.GetNames(typeof(FaceName))));
                        failCount++;
                        continue;
                    }

                    targetFaceB = SelectComponentFaceByRay(assyModel, compB, faceBKey);
                }

                try
                {
                    bool mateOk = AddCoincidentFaceMateFromResolvedFaces(swAssy, assyModel, compA, faceALabel, targetFaceA, compB, faceBLabel, targetFaceB, pair.IsFlipped);
                    if (mateOk)
                    {
                        Console.WriteLine("Face mate pair #" + (i + 1) + " succeeded: A " + faceALabel + " -> B " + faceBLabel + " (Flipped=" + pair.IsFlipped + ").");
                        successCount++;
                    }
                    else
                    {
                        Console.WriteLine("Face mate pair #" + (i + 1) + " failed: A " + faceALabel + " -> B " + faceBLabel + " (Flipped=" + pair.IsFlipped + ").");
                        failCount++;
                    }
                }
                catch (Exception pairEx)
                {
                    Console.WriteLine("Face mate pair #" + (i + 1) + " failed with runtime error: " + pairEx.Message);
                    failCount++;
                }
            }

            Console.WriteLine("Custom face mate summary: Success=" + successCount + ", Failed=" + failCount + ".");

            _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, oldInputDimVal);
            assyModel.ForceRebuild3(false);
            string assemblyPath = Path.Combine(folder, outputAssemblyFileName);
            assyModel.SaveAs3(assemblyPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent);
        }
        catch (Exception ex)
        {
            // Keep assembly run non-fatal when some mate operations fail.
            Console.WriteLine("Assembly generation warning: " + ex.Message);
        }
    }

    // Normalize user face names to canonical keys used by ray-selection logic.
    private static string NormalizeFaceName(string rawFaceName)
    {
        if (string.IsNullOrWhiteSpace(rawFaceName)) return "";

        string value = rawFaceName.Trim().ToLowerInvariant();
        switch (value)
        {
            case "top":
            case "up":
                return "top";
            case "bottom":
            case "down":
                return "bottom";
            case "left":
                return "left";
            case "right":
                return "right";
            case "front":
                return "front";
            case "back":
                return "back";
            default:
                return "";
        }
    }

    // Add one coincident mate between two component faces selected by geometric ray picking.
    private static bool AddCoincidentFaceMateByRay(AssemblyDoc swAssy, ModelDoc2 assyModel, Component2 compA, string faceA, Component2 compB, string faceB)
    {
        Face2 targetFaceA = SelectComponentFaceByRay(assyModel, compA, faceA);
        Face2 targetFaceB = SelectComponentFaceByRay(assyModel, compB, faceB);
        return AddCoincidentFaceMateFromResolvedFaces(swAssy, assyModel, compA, faceA, targetFaceA, compB, faceB, targetFaceB, false);
    }

    // Add one coincident mate between two already-resolved face entities.
    private static bool AddCoincidentFaceMateFromResolvedFaces(
        AssemblyDoc swAssy,
        ModelDoc2 assyModel,
        Component2 compA,
        string faceA,
        Face2 targetFaceA,
        Component2 compB,
        string faceB,
        Face2 targetFaceB,
        bool isFlipped)
    {
        if (targetFaceA == null || targetFaceB == null) return false;

        assyModel.ClearSelection2(true);
        bool selectedA = SelectFaceEntityForMate(assyModel, targetFaceA, false, 1);
        bool selectedB = SelectFaceEntityForMate(assyModel, targetFaceB, true, 1);
        if (!selectedA || !selectedB) return false;

        // Use requested mate alignment direction.
        int align = isFlipped
            ? (int)swMateAlign_e.swMateAlignANTI_ALIGNED
            : (int)swMateAlign_e.swMateAlignALIGNED;

        Feature createdMateFeature;
        bool createdAligned = TryCreateCoincidentMate(swAssy, assyModel, align, faceA, faceB, out createdMateFeature);
        if (!createdAligned)
        {
            Console.WriteLine("Mate creation failed for '" + faceA + "' -> '" + faceB + "' (Flipped=" + isFlipped + ").");
            // If a mate feature was created despite failure status, remove it for safety.
            if (createdMateFeature != null)
            {
                Console.WriteLine("Removing partial mate feature after failed attempt for '" + faceA + "' -> '" + faceB + "'.");
                DeleteFeatureIfExists(assyModel, createdMateFeature);
            }
            return false;
        }

        if (AreComponentsInterferingByBox(compA, compB))
        {
            Console.WriteLine("Interference detected after mate for '" + faceA + "' -> '" + faceB + "' (Flipped=" + isFlipped + "). Deleting mate and marking pair as failed.");
            DeleteFeatureIfExists(assyModel, createdMateFeature);
            return false;
        }

        return true;
    }

    // Create a coincident mate with requested alignment and return the created mate feature.
    private static bool TryCreateCoincidentMate(AssemblyDoc swAssy, ModelDoc2 assyModel, int align, string faceA, string faceB, out Feature createdFeature)
    {
        Feature before = GetLastFeature(assyModel);
        createdFeature = null;

        int mateError = 0;
        swAssy.AddMate3((int)swMateType_e.swMateCOINCIDENT, align, false, 0, 0, 0, 0, 0, 0, 0, 0, false, out mateError);

        Feature after = GetLastFeature(assyModel);
        if (after != null && !ReferenceEquals(before, after))
        {
            createdFeature = after;
        }

        // In some SOLIDWORKS states, AddMate3 can return a non-zero code even when a mate feature is created.
        // If a new feature exists, treat it as a successful mate to avoid unnecessary retry/disconnect paths.
        if (mateError != 0)
        {
            if (createdFeature != null)
            {
                Console.WriteLine("AddMate3 returned warning code: " + mateError + " but mate feature was created for faces '" + faceA + "' and '" + faceB + "'.");
                assyModel.ForceRebuild3(false);
                return true;
            }

            Console.WriteLine("AddMate3 returned error code: " + mateError + " for faces '" + faceA + "' and '" + faceB + "'.");
            return false;
        }

        assyModel.ForceRebuild3(false);

        return true;
    }

    // Approximate interference check using assembly-space component bounding boxes.
    private static bool AreComponentsInterferingByBox(Component2 compA, Component2 compB)
    {
        object boxAObj = compA.GetBox(false, false);
        object boxBObj = compB.GetBox(false, false);
        if (boxAObj == null || boxBObj == null) return false;

        double[] a = (double[])boxAObj;
        double[] b = (double[])boxBObj;
        if (a.Length < 6 || b.Length < 6) return false;

        double overlapX = Math.Min(a[3], b[3]) - Math.Max(a[0], b[0]);
        double overlapY = Math.Min(a[4], b[4]) - Math.Max(a[1], b[1]);
        double overlapZ = Math.Min(a[5], b[5]) - Math.Max(a[2], b[2]);

        const double tol = 1e-6;
        return overlapX > tol && overlapY > tol && overlapZ > tol;
    }

    // Return the last feature in the model feature tree.
    private static Feature GetLastFeature(ModelDoc2 model)
    {
        Feature current = model.FirstFeature();
        Feature last = null;
        while (current != null)
        {
            last = current;
            current = current.GetNextFeature();
        }

        return last;
    }

    // Delete a feature if available; used to remove wrong-direction mate before retry.
    private static void DeleteFeatureIfExists(ModelDoc2 model, Feature feat)
    {
        if (feat == null) return;
        bool selected = feat.Select2(false, 0);
        if (selected) model.EditDelete();
        model.ClearSelection2(true);
        model.ForceRebuild3(false);
    }

    // Select a target component face by casting rays and return matched planar Face2.
    private static Face2 SelectComponentFaceByRay(ModelDoc2 assyModel, Component2 comp, string faceKey)
    {
        object boxObj = comp.GetBox(false, false);
        if (boxObj == null) return null;

        double[] box = (double[])boxObj;
        if (box.Length < 6) return null;

        double minX = box[0], minY = box[1], minZ = box[2];
        double maxX = box[3], maxY = box[4], maxZ = box[5];

        double centerX = (minX + maxX) * 0.5;
        double centerY = (minY + maxY) * 0.5;
        double centerZ = (minZ + maxZ) * 0.5;

        double spanX = maxX - minX;
        double spanY = maxY - minY;
        double spanZ = maxZ - minZ;

        double margin = Math.Max(Math.Max(spanX, Math.Max(spanY, spanZ)) * 0.1, 0.002);
        double yProbe = centerY + (spanY * 0.22);

        double rayRadius = Math.Max(Math.Max(spanX, Math.Max(spanY, spanZ)) * 0.02, 0.0008);
        double xOffset = spanX * 0.18;
        double zOffset = spanZ * 0.18;
        double[][] rays;

        switch (faceKey.ToLowerInvariant())
        {
            case "top":
                rays = new[]
                {
                    new[] { centerX, yProbe, maxZ + margin, 0.0, 0.0, -1.0 },
                    new[] { centerX + xOffset, yProbe, maxZ + margin, 0.0, 0.0, -1.0 },
                    new[] { centerX - xOffset, yProbe, maxZ + margin, 0.0, 0.0, -1.0 }
                };
                break;
            case "bottom":
                rays = new[]
                {
                    new[] { centerX, yProbe, minZ - margin, 0.0, 0.0, 1.0 },
                    new[] { centerX + xOffset, yProbe, minZ - margin, 0.0, 0.0, 1.0 },
                    new[] { centerX - xOffset, yProbe, minZ - margin, 0.0, 0.0, 1.0 }
                };
                break;
            case "right":
                rays = new[]
                {
                    new[] { maxX + margin, yProbe, centerZ, -1.0, 0.0, 0.0 },
                    new[] { maxX + margin, yProbe, centerZ + zOffset, -1.0, 0.0, 0.0 },
                    new[] { maxX + margin, yProbe, centerZ - zOffset, -1.0, 0.0, 0.0 }
                };
                break;
            case "left":
                rays = new[]
                {
                    new[] { minX - margin, yProbe, centerZ, 1.0, 0.0, 0.0 },
                    new[] { minX - margin, yProbe, centerZ + zOffset, 1.0, 0.0, 0.0 },
                    new[] { minX - margin, yProbe, centerZ - zOffset, 1.0, 0.0, 0.0 }
                };
                break;
            case "front":
                rays = new[]
                {
                    new[] { centerX, maxY + margin, centerZ, 0.0, -1.0, 0.0 },
                    new[] { centerX + xOffset, maxY + margin, centerZ, 0.0, -1.0, 0.0 },
                    new[] { centerX - xOffset, maxY + margin, centerZ, 0.0, -1.0, 0.0 }
                };
                break;
            case "back":
                rays = new[]
                {
                    new[] { centerX, minY - margin, centerZ, 0.0, 1.0, 0.0 },
                    new[] { centerX + xOffset, minY - margin, centerZ, 0.0, 1.0, 0.0 },
                    new[] { centerX - xOffset, minY - margin, centerZ, 0.0, 1.0, 0.0 }
                };
                break;
            default:
                Console.WriteLine("Unknown face key: " + faceKey);
                return null;
        }

        SelectionMgr selMgr = (SelectionMgr)assyModel.SelectionManager;
        for (int i = 0; i < rays.Length; i++)
        {
            assyModel.ClearSelection2(true);
            bool hit = assyModel.Extension.SelectByRay(rays[i][0], rays[i][1], rays[i][2], rays[i][3], rays[i][4], rays[i][5], rayRadius, (int)swSelectType_e.swSelFACES, false, 0, 0);
            if (!hit) continue;

            object selectedObject = selMgr.GetSelectedObject6(1, -1);
            Face2 selectedFace = selectedObject as Face2;
            if (selectedFace == null) continue;
            if (!IsPlanarFace(selectedFace))
            {
                Console.WriteLine("Rejected non-planar face for '" + faceKey + "' on " + comp.Name2 + " (likely fillet).");
                continue;
            }

            Console.WriteLine("Ray face select '" + faceKey + "' on " + comp.Name2 + ": True (candidate " + i + ")");
            return selectedFace;
        }

        Console.WriteLine("Ray face select '" + faceKey + "' on " + comp.Name2 + ": False");
        return null;
    }

    // Check whether a face is planar; used to avoid selecting fillet/cylindrical faces.
    private static bool IsPlanarFace(Face2 face)
    {
        if (face == null) return false;
        Surface surface = (Surface)face.GetSurface();
        if (surface == null) return false;
        return surface.IsPlane();
    }

    // Select an exact face entity for mating with specific selection mark.
    private static bool SelectFaceEntityForMate(ModelDoc2 assyModel, Face2 face, bool append, int mark)
    {
        Entity entity = face as Entity;
        if (entity == null) return false;

        SelectionMgr selMgr = (SelectionMgr)assyModel.SelectionManager;
        SelectData selData = (SelectData)selMgr.CreateSelectData();
        selData.Mark = mark;
        bool selected = entity.Select4(append, selData);
        if (!selected)
        {
            selData.Mark = 0;
            selected = entity.Select4(append, selData);
        }

        return selected;
    }

    // Select a component face by point given in part-origin coordinates (mm).
    // Converts local part point -> assembly point using component transform, then selects by XYZ.
    private Face2 SelectComponentFaceByPartPoint(ModelDoc2 assyModel, Component2 comp, double xMm, double yMm, double zMm)
    {
        // Convert mm to meters for SOLIDWORKS API.
        double[] partPointMeters = { xMm / 1000.0, yMm / 1000.0, zMm / 1000.0 };

        // Get component transform from part space to assembly space.
        MathTransform compTransform = (MathTransform)comp.Transform2;
        if (compTransform == null)
        {
            Console.WriteLine("Point face select failed: component transform is null for " + comp.Name2 + ".");
            return null;
        }

        // Build math utility objects for transform multiplication.
        MathUtility mathUtil = (MathUtility)_swApp.GetMathUtility();
        if (mathUtil == null)
        {
            Console.WriteLine("Point face select failed: MathUtility is not available.");
            return null;
        }

        MathPoint localPoint = (MathPoint)mathUtil.CreatePoint(partPointMeters);
        MathPoint assemblyPoint = (MathPoint)localPoint.MultiplyTransform(compTransform);
        double[] p = (double[])assemblyPoint.ArrayData;

        // Select face using assembly-space XYZ.
        assyModel.ClearSelection2(true);
        bool selected = assyModel.Extension.SelectByID2("", "FACE", p[0], p[1], p[2], false, 1, null, 0);
        if (!selected)
        {
            selected = assyModel.Extension.SelectByID2("", "FACE", p[0], p[1], p[2], false, 0, null, 0);
        }

        if (!selected)
        {
            Console.WriteLine("Point face select failed at assembly point (" + p[0] + ", " + p[1] + ", " + p[2] + ") for " + comp.Name2 + ".");
            return null;
        }

        SelectionMgr selMgr = (SelectionMgr)assyModel.SelectionManager;
        Face2 selectedFace = selMgr.GetSelectedObject6(1, -1) as Face2;
        if (selectedFace == null)
        {
            Console.WriteLine("Point face select failed: selected object is not a face for " + comp.Name2 + ".");
            return null;
        }

        Console.WriteLine("Point face select succeeded for " + comp.Name2 + " at part point (" + xMm + ", " + yMm + ", " + zMm + ") mm.");
        return selectedFace;
    }

    // Select a named plane pair from component A and B in a robust order.
    private static bool SelectAssemblyPlanePair(ModelDoc2 assyModel, Component2 compA, Component2 compB, string assemblyName, string[] planeNames, int fallbackPlaneIndex, out bool statusA, out bool statusB)
    {
        statusA = false;
        statusB = false;
        string compAName = compA.Name2;
        string compBName = compB.Name2;

        string tokenA = ResolveAssemblyPlaneToken(assyModel, compA, assemblyName, planeNames, fallbackPlaneIndex);
        string tokenB = ResolveAssemblyPlaneToken(assyModel, compB, assemblyName, planeNames, fallbackPlaneIndex);
        statusA = !string.IsNullOrWhiteSpace(tokenA);
        statusB = !string.IsNullOrWhiteSpace(tokenB);
        if (!statusA || !statusB) return false;

        assyModel.ClearSelection2(true);
        bool selectedA = SelectAssemblyPlaneToken(assyModel, tokenA, compAName, assemblyName, false);
        bool selectedB = SelectAssemblyPlaneToken(assyModel, tokenB, compBName, assemblyName, true);
        if (selectedA && selectedB)
        {
            statusA = true;
            statusB = true;
            return true;
        }

        assyModel.ClearSelection2(true);
        selectedB = SelectAssemblyPlaneToken(assyModel, tokenB, compBName, assemblyName, false);
        selectedA = SelectAssemblyPlaneToken(assyModel, tokenA, compAName, assemblyName, true);
        statusA = selectedA;
        statusB = selectedB;
        return selectedA && selectedB;
    }

    // Resolve a usable assembly-plane token by preferred names then feature-tree fallback index.
    private static string ResolveAssemblyPlaneToken(ModelDoc2 assyModel, Component2 component, string assemblyName, string[] preferredPlaneNames, int fallbackPlaneIndex)
    {
        string componentName = component.Name2;
        for (int i = 0; i < preferredPlaneNames.Length; i++)
        {
            if (SelectAssemblyPlaneToken(assyModel, preferredPlaneNames[i], componentName, assemblyName, false))
            {
                Console.WriteLine("Assembly plane resolved by name '" + preferredPlaneNames[i] + "' for component '" + componentName + "'.");
                assyModel.ClearSelection2(true);
                return preferredPlaneNames[i];
            }
        }

        ModelDoc2 componentModel = (ModelDoc2)component.GetModelDoc2();
        string fallbackPlaneName = GetReferencePlaneNameByIndex(componentModel, fallbackPlaneIndex);
        if (!string.IsNullOrWhiteSpace(fallbackPlaneName))
        {
            if (SelectAssemblyPlaneToken(assyModel, fallbackPlaneName, componentName, assemblyName, false))
            {
                Console.WriteLine("Assembly plane resolved by fallback feature scan: '" + fallbackPlaneName + "' for component '" + componentName + "'.");
                assyModel.ClearSelection2(true);
                return fallbackPlaneName;
            }

            Console.WriteLine("Assembly fallback plane found but could not be selected: '" + fallbackPlaneName + "' for component '" + componentName + "'.");
            assyModel.ClearSelection2(true);
            return "";
        }

        Console.WriteLine("Assembly fallback plane not found at index " + fallbackPlaneIndex + " for component '" + componentName + "'.");
        return "";
    }

    // Try multiple plane selection token formats for assembly context compatibility.
    private static bool SelectAssemblyPlaneToken(ModelDoc2 assyModel, string planeName, string componentName, string assemblyName, bool append)
    {
        string[] selections =
        {
            planeName + "@" + componentName + "@" + assemblyName,
            planeName + "@" + componentName
        };

        for (int i = 0; i < selections.Length; i++)
        {
            bool selected = assyModel.Extension.SelectByID2(selections[i], "PLANE", 0, 0, 0, append, 1, null, 0);
            if (!selected)
            {
                selected = assyModel.Extension.SelectByID2(selections[i], "PLANE", 0, 0, 0, append, 0, null, 0);
            }

            if (selected) return true;
        }

        return false;
    }

    // Get the Nth reference plane name from the component feature tree.
    private static string GetReferencePlaneNameByIndex(ModelDoc2 model, int planeIndex)
    {
        if (model == null || planeIndex < 0) return "";

        int currentIndex = 0;
        Feature feat = model.FirstFeature();
        while (feat != null)
        {
            string featureType = feat.GetTypeName2();
            if (string.Equals(featureType, "RefPlane", StringComparison.OrdinalIgnoreCase))
            {
                if (currentIndex == planeIndex) return feat.Name;
                currentIndex++;
            }

            feat = feat.GetNextFeature();
        }

        return "";
    }

    // Print reference plane names for one component for diagnostics.
    private static void LogComponentReferencePlanes(Component2 component, string label)
    {
        ModelDoc2 componentModel = (ModelDoc2)component.GetModelDoc2();
        if (componentModel == null)
        {
            Console.WriteLine("Component " + label + " has no loaded model for plane logging.");
            return;
        }

        Console.WriteLine("Component " + label + " reference planes:");
        int planeCount = 0;
        Feature feat = componentModel.FirstFeature();
        while (feat != null)
        {
            if (string.Equals(feat.GetTypeName2(), "RefPlane", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("  [" + planeCount + "] " + feat.Name);
                planeCount++;
            }

            feat = feat.GetNextFeature();
        }

        if (planeCount == 0)
        {
            Console.WriteLine("  (none found)");
        }
    }
}
