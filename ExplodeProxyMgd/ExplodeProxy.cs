using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.ApplicationServices;

namespace ExplodeProxyMgd
{
    public class ExplodeProxy
    {
        
        /// <summary>
        /// Explode proxy entity to blockreference
        /// Return string name of blockreference in drawing database
        /// </summary>
        /// <param name="rbArgs">Lisp arguments: entity objectId, string BlockPrefix</param>
        /// <returns>string BlockReference name</returns>
        [LispFunction("proxy-explode-to-block")]
        public static TypedValue ProxyExplodeToBlock(ResultBuffer rbArgs)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            TypedValue res = new TypedValue((int)LispDataType.Text,"");

            if (rbArgs.AsArray().Length == 2)
            {
                TypedValue entity = rbArgs.AsArray()[0];
                TypedValue blkPrefix = rbArgs.AsArray()[1];
                
                if ((entity.TypeCode == (int)LispDataType.ObjectId) && (blkPrefix.TypeCode == (int)LispDataType.Text))
                {
                    using (Transaction tr = doc.TransactionManager.StartTransaction())
                    {
                        try
                        {
                            ObjectId id = (ObjectId)entity.Value;
                            DBObjectCollection objs = new DBObjectCollection();
                            BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                            Entity entx = (Entity)tr.GetObject(id, OpenMode.ForWrite);
                            entx.Explode(objs);

                            string blkName = blkPrefix.Value.ToString() + entx.Handle.ToString();

                            if (bt.Has(blkName) == false)
                            {
                                BlockTableRecord btr = new BlockTableRecord();
                                btr.Name = blkName;

                                bt.UpgradeOpen();
                                ObjectId btrId = bt.Add(btr);
                                tr.AddNewlyCreatedDBObject(btr, true);

                                foreach (DBObject obj in objs)
                                {
                                    Entity ent = (Entity)obj;
                                    btr.AppendEntity(ent);
                                    tr.AddNewlyCreatedDBObject(ent, true);
                                }
                            }
                            res = new TypedValue((int)LispDataType.Text, blkName);

                            tr.Commit();
                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception ex)
                        {
                            tr.Abort();
                            ed.WriteMessage(ex.Message);
                        }
                    }
                }
            }
            return res;
        }

        /// <summary>
        /// Explode proxy entity in place
        /// </summary>
        /// <param name="rbArgs">entity name</param>
        [LispFunction("proxy-explode-in-place")]
        public static void ProxyExplodeInPlace(ResultBuffer rbArgs)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            if (rbArgs.AsArray().Length == 1)
            {
                TypedValue entity = rbArgs.AsArray().First();

                if (entity.TypeCode == (int)LispDataType.ObjectId)
                {
                    using (Transaction tr = doc.TransactionManager.StartTransaction())
                    {
                        try
                        {
                            ObjectId id = (ObjectId)entity.Value;
                            DBObjectCollection objs = new DBObjectCollection();
                            BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                            Entity entx = (Entity)tr.GetObject(id, OpenMode.ForWrite);
                            entx.Explode(objs);

                            entx.Erase();

                            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                            foreach (DBObject obj in objs)
                            {
                                Entity ent = (Entity)obj;
                                btr.AppendEntity(ent);
                                tr.AddNewlyCreatedDBObject(ent, true);
                            }

                            tr.Commit();
                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception ex)
                        {
                            tr.Abort();
                            ed.WriteMessage(ex.Message);
                        }                        
                    }
                }
            }
        }
    }
}
