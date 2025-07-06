// AutoCAD utility plugin with multiple custom commands
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

[assembly: CommandClass(typeof(MyAutoCADPlugin.UtilityCommands))]

namespace MyAutoCADPlugin
{
    public class UtilityCommands
    {
        [CommandMethod("MeasureLineSimple")]
        public void MeasureLineSimple()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            PromptEntityOptions peo = new PromptEntityOptions("\nSelect a line: ");
            peo.SetRejectMessage("\nOnly lines allowed.");
            peo.AddAllowedClass(typeof(Line), exactMatch: true);

            PromptEntityResult per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK)
                return;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                Line line = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Line;
                if (line != null)
                {
                    ed.WriteMessage($"\nLength: {line.Length:F2} units");
                }
                tr.Commit();
            }
        }


        [CommandMethod("ADDLABEL")]
        public void AddLabelToBlocks()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            // Select blocks to label
            PromptSelectionOptions selOpts = new PromptSelectionOptions();
            selOpts.MessageForAdding = "\nSelect blocks to add labels:";
            SelectionFilter filter = new SelectionFilter(new[] { new TypedValue((int)DxfCode.Start, "INSERT") });
            PromptSelectionResult selRes = ed.GetSelection(selOpts, filter);
            if (selRes.Status != PromptStatus.OK) return;

            // Get prefix
            PromptStringOptions pso = new PromptStringOptions("\nEnter label prefix (e.g. BLK): ");
            pso.AllowSpaces = false;
            PromptResult prefixRes = ed.GetString(pso);
            if (prefixRes.Status != PromptStatus.OK) return;
            string prefix = prefixRes.StringResult;

            // Get text height
            PromptDoubleOptions heightOpts = new PromptDoubleOptions("\nEnter text height: ");
            heightOpts.AllowNegative = false;
            heightOpts.AllowZero = false;
            heightOpts.DefaultValue = 10.0; // Better default
            PromptDoubleResult heightRes = ed.GetDouble(heightOpts);
            if (heightRes.Status != PromptStatus.OK) return;
            double textHeight = heightRes.Value;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // Open Block Table and Model Space (write mode)
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                // Get the TextStyleTable and the 'Standard' style
                TextStyleTable tst = (TextStyleTable)tr.GetObject(db.TextStyleTableId, OpenMode.ForRead);
                ObjectId textStyleId = tst.Has("Standard") ? tst["Standard"] : ObjectId.Null;

                int i = 1;
                foreach (ObjectId blkId in selRes.Value.GetObjectIds())
                {
                    BlockReference blkRef = tr.GetObject(blkId, OpenMode.ForRead) as BlockReference;
                    if (blkRef == null) continue;

                    // Position label above the block reference
                    Point3d labelPos = blkRef.Position + new Vector3d(0, textHeight * 1.5, 0);

                    DBText label = new DBText
                    {
                        Position = labelPos,
                        Height = textHeight,
                        TextString = $"{prefix}{i++}",
                        Color = Color.FromColorIndex(ColorMethod.ByAci, 3), // Green
                        TextStyleId = textStyleId,
                        Layer = blkRef.Layer // Optional: place on same layer as block
                    };

                    ms.AppendEntity(label);
                    tr.AddNewlyCreatedDBObject(label, true);
                }

                tr.Commit();
            }
            ed.WriteMessage($"\nAdded labels with prefix '{prefix}' to {selRes.Value.Count} blocks.");
        }

        [CommandMethod("ARRAYBLOCKLINE")]
        public void ArrayBlockOnLineImproved()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            // Get block name from user
            var blockNameOpts = new PromptStringOptions("\nEnter block name: ") { AllowSpaces = false };
            var blockNameRes = ed.GetString(blockNameOpts);
            if (blockNameRes.Status != PromptStatus.OK) return;
            string blockName = blockNameRes.StringResult;

            ObjectId blockId = ObjectId.Null;
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                if (!bt.Has(blockName))
                {
                    ed.WriteMessage($"\nBlock '{blockName}' not found.");
                    return;
                }
                blockId = bt[blockName];
                tr.Commit();
            }

            // Select line/polyline
            var lineOpts = new PromptEntityOptions("\nSelect line to array blocks along: ");
            lineOpts.SetRejectMessage("\nMust be a line or polyline.");
            lineOpts.AddAllowedClass(typeof(Line), true);
            lineOpts.AddAllowedClass(typeof(Polyline), false);
            var lineRes = ed.GetEntity(lineOpts);
            if (lineRes.Status != PromptStatus.OK) return;

            // Choose method
            var methodOpts = new PromptKeywordOptions("\nArray by [Count/Distance]: ", "Count Distance") { Keywords = { Default = "Count" } };
            var methodRes = ed.GetKeywords(methodOpts);
            if (methodRes.Status != PromptStatus.OK) return;
            bool useCount = methodRes.StringResult == "Count";

            int count = 0;
            double spacing = 0;

            if (useCount)
            {
                var countOpts = new PromptIntegerOptions("\nNumber of blocks: ")
                {
                    AllowNegative = false,
                    AllowZero = false,
                    DefaultValue = 5
                };
                var countRes = ed.GetInteger(countOpts);
                if (countRes.Status != PromptStatus.OK) return;
                count = countRes.Value;
            }
            else
            {
                var distOpts = new PromptDoubleOptions("\nEnter spacing between blocks: ")
                {
                    AllowNegative = false,
                    AllowZero = false,
                    DefaultValue = 10.0
                };
                var distRes = ed.GetDouble(distOpts);
                if (distRes.Status != PromptStatus.OK) return;
                spacing = distRes.Value;
            }

            var scaleOpts = new PromptDoubleOptions("\nEnter block scale factor [1.0]: ")
            {
                AllowNegative = false,
                AllowZero = false,
                DefaultValue = 1.0,
                UseDefaultValue = true
            };
            var scaleRes = ed.GetDouble(scaleOpts);
            if (scaleRes.Status != PromptStatus.OK) return;
            double scaleFactor = scaleRes.Value;

            var rotOpts = new PromptKeywordOptions("\nAlign blocks with line? [Yes/No]: ", "Yes No") { Keywords = { Default = "Yes" } };
            var rotRes = ed.GetKeywords(rotOpts);
            if (rotRes.Status != PromptStatus.OK) return;
            bool alignWithLine = rotRes.StringResult == "Yes";

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var ent = tr.GetObject(lineRes.ObjectId, OpenMode.ForRead);
                Point3d startPoint, endPoint;
                Vector3d direction;
                double totalLength;

                if (ent is Line line)
                {
                    startPoint = line.StartPoint;
                    endPoint = line.EndPoint;
                    direction = (endPoint - startPoint).GetNormal();
                    totalLength = line.Length;
                }
                else if (ent is Polyline pl && pl.NumberOfVertices >= 2)
                {
                    startPoint = pl.GetPoint3dAt(0);
                    endPoint = pl.GetPoint3dAt(pl.NumberOfVertices - 1);
                    direction = (endPoint - startPoint).GetNormal();
                    totalLength = pl.Length;
                }
                else
                {
                    ed.WriteMessage("\nUnsupported entity.");
                    return;
                }

                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                var blockDef = (BlockTableRecord)tr.GetObject(blockId, OpenMode.ForRead);

                // Get block center offset (important!)
                Vector3d centerOffset = new Vector3d(0, 0, 0);
                if (blockDef.Bounds.HasValue)
                {
                    var bounds = blockDef.Bounds.Value;
                    var center = bounds.MinPoint + ((bounds.MaxPoint - bounds.MinPoint) / 2.0);
                    centerOffset = center.GetAsVector();
                }

                // Calculate arraying parameters
                int numBlocks;
                double actualSpacing;

                if (useCount)
                {
                    numBlocks = count;
                    actualSpacing = count > 1 ? totalLength / (count - 1) : 0;
                }
                else
                {
                    numBlocks = (int)Math.Floor(totalLength / spacing) + 1;
                    actualSpacing = spacing;
                }

                double rotationAngle = alignWithLine ? Math.Atan2(direction.Y, direction.X) : 0;
                int blocksCreated = 0;

                for (int i = 0; i < numBlocks; i++)
                {
                    Point3d insertPoint = startPoint + direction * actualSpacing * i;

                    if (insertPoint.DistanceTo(startPoint) > totalLength + 0.001)
                        break;

                    var br = new BlockReference(insertPoint, blockId)
                    {
                        Rotation = rotationAngle,
                        ScaleFactors = new Scale3d(scaleFactor)
                    };

                    // Move block reference to center
                    Matrix3d moveToCenter = Matrix3d.Displacement(-centerOffset);
                    br.TransformBy(moveToCenter);

                    ms.AppendEntity(br);
                    tr.AddNewlyCreatedDBObject(br, true);

                    // Attributes
                    if (blockDef.HasAttributeDefinitions)
                    {
                        foreach (ObjectId attId in blockDef)
                        {
                            var obj = tr.GetObject(attId, OpenMode.ForRead);
                            if (obj is AttributeDefinition attDef && !attDef.Constant)
                            {
                                var attRef = new AttributeReference();
                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);

                                if (attDef.Tag.ToUpper().Contains("NUM") || attDef.Tag.ToUpper().Contains("ID"))
                                {
                                    attRef.TextString = (i + 1).ToString("D2");
                                }
                                else
                                {
                                    attRef.TextString = attDef.TextString;
                                }
                                
                                br.AttributeCollection.AppendAttribute(attRef);
                                tr.AddNewlyCreatedDBObject(attRef, true);
                            }
                        }
                    }

                    blocksCreated++;
                }

                tr.Commit();
                ed.WriteMessage($"\nCreated {blocksCreated} '{blockName}' blocks on the line.");
                if (useCount && blocksCreated < count)
                {
                    ed.WriteMessage($"\nNote: Only {blocksCreated} blocks fit on the line.");

                }
            }
        }



    }
}