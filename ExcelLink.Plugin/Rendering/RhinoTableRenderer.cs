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
        bool insertInstance,
        out int instanceDefinitionIndex,
        out Guid insertedInstanceId,
        System.Collections.Generic.IEnumerable<Transform>? reinsertTransforms = null)
    {
        instanceDefinitionIndex = -1;
        insertedInstanceId = Guid.Empty;
        if (table.Columns.Count == 0 || table.Rows.Count == 0)
            return Result.Failure;

        double mmToModel = RhinoMath.UnitScale(UnitSystem.Millimeters, doc.ModelUnitSystem) * Math.Max(1e-6, scaleMultiplier);
        double textHeightMm = 2.5;
        double marginMm = 1.0;
        double textHeight = textHeightMm * mmToModel;
        double margin = marginMm * mmToModel;

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

        double totalHeight = yPositions[^1];

        // Grid - horizontal (invert so first row appears at top)
        for (int r = 0; r < yPositions.Count; r++)
        {
            double y = topRowAtTop ? (totalHeight - yPositions[r]) : yPositions[r];
            var a = new Point3d(0, y, 0);
            var b = new Point3d(xPositions[^1], y, 0);
            geoms.Add(new LineCurve(a, b));
            atts.Add(new ObjectAttributes());
        }

        // Grid - vertical
        for (int c = 0; c < xPositions.Count; c++)
        {
            var a = new Point3d(xPositions[c], 0, 0);
            var b = new Point3d(xPositions[c], yPositions[^1], 0);
            geoms.Add(new LineCurve(a, b));
            atts.Add(new ObjectAttributes());
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

                var te = new TextEntity
                {
                    Plane = plane,
                    PlainText = text,
                    TextHeight = textHeight,
                };

                te.Justification = effective switch
                {
                    HorizontalAlignment.Center => TextJustification.MiddleCenter,
                    HorizontalAlignment.Right => TextJustification.MiddleRight,
                    _ => TextJustification.MiddleLeft
                };

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

        var basePlane = Plane.WorldXY;
        int idefIndex = idefTable.Add(blockName, string.Empty, basePlane.Origin, geoms, atts);
        if (idefIndex < 0)
            return Result.Failure;

        var def = idefTable.Find(blockName);
        if (def == null)
            return Result.Failure;

        instanceDefinitionIndex = def.Index;
        if (reinsertTransforms != null)
        {
            foreach (var xf in reinsertTransforms)
                doc.Objects.AddInstanceObject(def.Index, xf);
        }
        else if (insertInstance)
        {
            var xform = Transform.Translation(insertPoint - Point3d.Origin);
            insertedInstanceId = doc.Objects.AddInstanceObject(def.Index, xform);
        }
        doc.Views.Redraw();
        return Result.Success;
    }
}


