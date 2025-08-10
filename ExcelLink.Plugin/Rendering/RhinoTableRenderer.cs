using System;
using System.Collections.Generic;
using Rhino;
using Rhino.DocObjects;
using Rhino.Geometry;
using Rhino.Commands;
using Rhino.Collections;
using ExcelLink.Core.Models;

namespace ExcelLink.Plugin.Rendering;

public static class RhinoTableRenderer
{
    public enum HorizontalOverride
    {
        UseExcel,
        Left,
        Center,
        Right
    }

    public static Result RenderTableAsBlock(
        RhinoDoc doc,
        TableModel table,
        string blockName,
        Point3d insertPoint,
        double scaleMultiplier,
        bool topRowAtTop,
        HorizontalOverride horizontal,
        int verticalAlign, // 0=Top,1=Middle,2=Bottom
        bool insertInstance,
        out int instanceDefinitionIndex,
        out Guid insertedInstanceId,
        out System.Collections.Generic.List<Guid>? reinsertedInstanceIds,
        System.Collections.Generic.IEnumerable<Transform>? reinsertTransforms = null,
        bool drawGrid = true,
        double? overrideTextHeightMm = null,
        string? fontFamily = null,
        string? dimStyleName = null,
        bool enableWrap = false)
    {
        instanceDefinitionIndex = -1;
        insertedInstanceId = Guid.Empty;
        reinsertedInstanceIds = null;
        if (table.Columns.Count == 0 || table.Rows.Count == 0)
            return Result.Failure;

        double mmToModel = RhinoMath.UnitScale(UnitSystem.Millimeters, doc.ModelUnitSystem) * Math.Max(1e-6, scaleMultiplier);
        double textHeightMm = overrideTextHeightMm ?? 2.5;
        double marginMm = 1.0;
        double textHeight = textHeightMm * mmToModel;
        double margin = marginMm * mmToModel;

        // Resolve DimStyle once if provided
        Guid? dimStyleId = null;
        int dimStyleIndex = -1;
        if (!string.IsNullOrWhiteSpace(dimStyleName))
        {
            var ds = doc.DimStyles.FindName(dimStyleName);
            if (ds != null)
            {
                dimStyleId = ds.Id;
                dimStyleIndex = ds.Index;
            }
        }

        // Precompute x and y coordinates of grid lines
        var xPositions = new List<double> { 0.0 };
        foreach (var col in table.Columns)
        {
            double widthModel = (col.WidthMillimeters <= 0 ? 10.0 : col.WidthMillimeters) * mmToModel;
            xPositions.Add(xPositions[^1] + widthModel);
        }

        var yPositions = new List<double> { 0.0 };
        foreach (var row in table.Rows)
        {
            double heightModel = (row.HeightMillimeters <= 0 ? 5.0 : row.HeightMillimeters) * mmToModel;
            yPositions.Add(yPositions[^1] + heightModel);
        }

        var geoms = new List<GeometryBase>();
        var atts = new List<ObjectAttributes>();
        var tempIds = new List<Guid>();

        double totalHeight = yPositions[^1];

        // Grid - horizontal (invert so first row appears at top)
        if (drawGrid)
        {
            for (int r = 0; r < yPositions.Count; r++)
            {
                double y = topRowAtTop ? (totalHeight - yPositions[r]) : yPositions[r];
                var a = new Point3d(0, y, 0);
                var b = new Point3d(xPositions[^1], y, 0);
                geoms.Add(new LineCurve(a, b));
                atts.Add(new ObjectAttributes());
            }
        }

        // Grid - vertical
        if (drawGrid)
        {
            for (int c = 0; c < xPositions.Count; c++)
            {
                var a = new Point3d(xPositions[c], 0, 0);
                var b = new Point3d(xPositions[c], yPositions[^1], 0);
                geoms.Add(new LineCurve(a, b));
                atts.Add(new ObjectAttributes());
            }
        }

        // Text
        for (int r = 0; r < table.Rows.Count; r++)
        {
            var row = table.Rows[r];
            for (int c = 0; c < table.Columns.Count && c < row.Cells.Count; c++)
            {
                var cell = row.Cells[c];
                string text = cell.Value?.ToString() ?? string.Empty;
                if (string.IsNullOrEmpty(text))
                    continue;

                double x0 = xPositions[c] + margin;
                double x1 = xPositions[c + 1] - margin;
                double rowTop = topRowAtTop ? (totalHeight - yPositions[r]) : yPositions[r];
                double rowBottom = topRowAtTop ? (totalHeight - yPositions[r + 1]) : yPositions[r + 1];
                double y0 = Math.Min(rowTop, rowBottom) + margin;
                double y1 = Math.Max(rowTop, rowBottom) - margin;
                // Determine anchor based on alignment
                var effective = horizontal switch
                {
                    HorizontalOverride.Left => HorizontalAlignment.Left,
                    HorizontalOverride.Center => HorizontalAlignment.Center,
                    HorizontalOverride.Right => HorizontalAlignment.Right,
                    _ => cell.HorizontalAlignment
                };

                Point3d anchor;
                switch (effective)
                {
                    case HorizontalAlignment.Left:
                        anchor = new Point3d(x0, (y0 + y1) * 0.5, 0);
                        break;
                    case HorizontalAlignment.Right:
                        anchor = new Point3d(x1, (y0 + y1) * 0.5, 0);
                        break;
                    default:
                        anchor = new Point3d((x0 + x1) * 0.5, (y0 + y1) * 0.5, 0);
                        break;
                }

                var plane = Plane.WorldXY;
                plane.Origin = anchor;

                // optional wrap: naive soft wrap by character count based on column width
                if (enableWrap)
                {
                    double approxCharWidth = textHeight * 0.6; // heuristic
                    double maxWidth = Math.Max(0.0, x1 - x0);
                    int maxChars = (int)Math.Max(1, Math.Floor(maxWidth / Math.Max(1e-6, approxCharWidth)));
                    if (maxChars > 0 && text.Length > maxChars)
                    {
                        text = WrapText(text, maxChars);
                    }
                }

                TextEntity te;
                bool appliedDimStyle = false;
                if (dimStyleId.HasValue)
                {
                    // Use API factory that binds a DimensionStyle at creation time
                    var ds = doc.DimStyles.FindId(dimStyleId.Value);
                    bool wrapped = enableWrap;
                    double rectWidth = enableWrap ? Math.Max(1e-6, x1 - x0) : 0.0;
                    te = TextEntity.Create(text, plane, ds, wrapped, rectWidth, 0.0);
                    appliedDimStyle = true;
                }
                else
                {
                    te = new TextEntity { Plane = plane, PlainText = text };
                    if (!string.IsNullOrWhiteSpace(fontFamily))
                    {
                        te.Font = Font.FromQuartetProperties(fontFamily, false, false);
                    }
                }

                // Only set explicit text height when
                // - caller provided an override OR
                // - no DimStyle applied (keep DimStyle height intact)
                if (overrideTextHeightMm.HasValue || !appliedDimStyle)
                {
                    te.TextHeight = textHeight;
                }

                var baseJust = effective switch
                {
                    HorizontalAlignment.Center => TextJustification.MiddleCenter,
                    HorizontalAlignment.Right => TextJustification.MiddleRight,
                    _ => TextJustification.MiddleLeft
                };
                // Map vertical alignment into TextJustification quadrant
                te.Justification = baseJust;
                if (verticalAlign == 0)
                {
                    if (baseJust == TextJustification.MiddleLeft) te.Justification = TextJustification.TopLeft;
                    else if (baseJust == TextJustification.MiddleCenter) te.Justification = TextJustification.TopCenter;
                    else if (baseJust == TextJustification.MiddleRight) te.Justification = TextJustification.TopRight;
                }
                else if (verticalAlign == 2)
                {
                    if (baseJust == TextJustification.MiddleLeft) te.Justification = TextJustification.BottomLeft;
                    else if (baseJust == TextJustification.MiddleCenter) te.Justification = TextJustification.BottomCenter;
                    else if (baseJust == TextJustification.MiddleRight) te.Justification = TextJustification.BottomRight;
                }

                geoms.Add(te);
                atts.Add(new ObjectAttributes());
            }
        }

        // Create or replace block definition
        var idefTable = doc.InstanceDefinitions;
        var existing = idefTable.Find(blockName);
        if (existing != null)
        {
            idefTable.Delete(existing.Index, true, true);
        }

        // Add temporary objects to document to preserve annotation styles reliably
        for (int i = 0; i < geoms.Count; i++)
        {
            var g = geoms[i];
            var a = atts[i];
            Guid id;
            switch (g)
            {
                case LineCurve lc:
                    id = doc.Objects.AddCurve(lc, a);
                    break;
                case TextEntity te:
                    id = doc.Objects.AddText(te, a);
                    break;
                default:
                    id = doc.Objects.Add(g, a);
                    break;
            }
            if (id != Guid.Empty) tempIds.Add(id);
        }

        var basePlane = Plane.WorldXY;
        // Retrieve geometries from ids for definition
        var objs = tempIds.Select(id => doc.Objects.FindId(id)).Where(o => o != null).ToList();
        var defGeoms = new List<GeometryBase>();
        var defAtts = new List<ObjectAttributes>();
        foreach (var o in objs)
        {
            defGeoms.Add(o.Geometry.Duplicate());
            defAtts.Add(o.Attributes.Duplicate());
        }
        int idefIndex = idefTable.Add(blockName, string.Empty, basePlane.Origin, defGeoms, defAtts);
        if (idefIndex < 0)
            return Result.Failure;

        var def = idefTable.Find(blockName);
        if (def == null)
            return Result.Failure;

        // cleanup temp objects
        foreach (var id in tempIds)
            doc.Objects.Delete(id, true);

        instanceDefinitionIndex = def.Index;
        if (reinsertTransforms != null)
        {
            reinsertedInstanceIds = new List<Guid>();
            foreach (var xf in reinsertTransforms)
            {
                var id = doc.Objects.AddInstanceObject(def.Index, xf);
                if (id != Guid.Empty) reinsertedInstanceIds.Add(id);
            }
        }
        else if (insertInstance)
        {
            var xform = Transform.Translation(insertPoint - Point3d.Origin);
            insertedInstanceId = doc.Objects.AddInstanceObject(def.Index, xform);
        }
        doc.Views.Redraw();
        return Result.Success;
    }

    private static string WrapText(string input, int maxChars)
    {
        var words = input.Split(' ');
        var line = string.Empty;
        var lines = new List<string>();
        foreach (var w in words)
        {
            if ((line.Length + (line.Length > 0 ? 1 : 0) + w.Length) > maxChars)
            {
                if (!string.IsNullOrEmpty(line)) lines.Add(line);
                line = w;
            }
            else
            {
                line = string.IsNullOrEmpty(line) ? w : line + " " + w;
            }
        }
        if (!string.IsNullOrEmpty(line)) lines.Add(line);
        return string.Join("\n", lines);
    }
}


