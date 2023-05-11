using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using System.IO;
using System.Linq;
using Microsoft.Win32;
using Autodesk.AutoCAD.Windows;

namespace PurchasedParts
{
    //requests a file path to use to insert the drawing as a block
    //macro should be gives name of block
    //looks through the registry to find a key with the name of block
    //if key doesnt exist asks for a file path to use and loads
    //if key exists trys to get the path from its values
    //if path isnt valid, it will try that path altered to the available drives
    //if all drives fail, it will request a file path from user to use for future ref
    //follows cursor for  insert point

        //TODO
        //blocks respect UCS but not dynamic UCS


    //Jig to be let user move block around during insertion
    class BlockJig: EntityJig
    {
        #region fields
        Point3d insertPoint, mActualPoint;
        CoordinateSystem3d insert_Ucs, current_Ucs;
        #endregion

        #region Constructors
        public BlockJig(BlockReference br)
            :base(br)
        { insertPoint = br.Position;
            br.TransformBy(Application.DocumentManager.MdiActiveDocument.Editor.CurrentUserCoordinateSystem);
        }

        #endregion

        #region Properties
        //need editor to get the ucs on the fly
        private Editor Editor
        { get { return Application.DocumentManager.MdiActiveDocument.Editor; } }

        private Matrix3d UCS
        { get { return Editor.CurrentUserCoordinateSystem; } }

        private CoordinateSystem3d cs
        { get { return UCS.CoordinateSystem3d; } }
        #endregion

        #region Overrides
        protected override SamplerStatus Sampler(JigPrompts prompts)
        {
            JigPromptPointOptions jigOpts = new JigPromptPointOptions();
            jigOpts.UserInputControls = (UserInputControls.Accept3dCoordinates)
                | UserInputControls.NoZeroResponseAccepted
                | UserInputControls.NoNegativeResponseAccepted;
            jigOpts.Message = "\nEnter insert point: ";

            PromptPointResult ppr = prompts.AcquirePoint(jigOpts);

            if (mActualPoint == ppr.Value)
                return SamplerStatus.NoChange;
            else
                mActualPoint = ppr.Value;
                //insert_Ucs = prompts.               

            return SamplerStatus.OK;
        }

        protected override bool Update()
        {
            insertPoint = mActualPoint;

            ////transform from ucs to wcs
            //Matrix3d mat =
            //    Matrix3d.AlignCoordinateSystem(
            //        Point3d.Origin,
            //        Vector3d.XAxis,
            //        Vector3d.YAxis,
            //        Vector3d.ZAxis,
            //        cs.Origin,
            //        cs.Xaxis,
            //        cs.Yaxis,
            //        cs.Zaxis
            //        );

            //insertPoint = insertPoint.TransformBy(mat.Inverse());

            //Matrix3d mat = Matrix3d.Displacement(insertPoint.GetVectorTo(mActualPoint));
            //Matrix3d mat = Matrix3d.WorldToPlane(entitry.vector);
            //point = firstPoint.TransformBy(mat);

            //Vector3d xdir = (Point3d)Application.GetSystemVariable("UCSXDIR") - Point3d.Origin;
            //double ucsRot = Vector3d.XAxis.GetAngleTo(xdir, Vector3d.ZAxis);

            Matrix3d wcs2ucs = UCS.Inverse();

            try
            {
                //insertPoint = insertPoint.TransformBy(wcs2ucs);
                ((BlockReference)Entity).Position = insertPoint;

                //apply matrix to rotate part
                //Entity.TransformBy(mat);

                //Entity.TransformBy(Matrix3d.Rotation(ucsRot, Vector3d.ZAxis, insertPoint));
            }
            catch (System.Exception)
            { return false; }
            
            return true;
        }
        #endregion

        public Entity GetEntity()
        { return Entity; }
    }

    //command class
    public class Class1
    {
        [CommandMethod("QuickInsert")]
        public void QuickInsert()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            bool updatePath = false;
            string filePath = null;
            string blockName;
            PromptStringOptions pso = new PromptStringOptions("Block to insert?");
            pso.AllowSpaces = true;
            PromptResult res = ed.GetString(pso);
            if (res.Status == PromptStatus.OK)
                blockName = res.StringResult;
            else
                return;

            //shouldn't ever be a problem because macros should always have a name
            if (blockName == null | blockName == "")
                return;

            #region Get FilePath from registry
            //use the block name requested to look up a file path saved in a registry key
            //first check to see if that key exists
            string regKey = @"Software\Autodesk\AutoCAD\" + blockName;
            Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(regKey, true);
            if (key == null)
            {
                //Key doesnt exist
                //create a key
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(regKey);
                //create a path to ref in the future as key value
                updatePath = true;
                filePath = pickSpec(ed, doc);
            }
            else
            {
                //key exists
                //get key value and verify it isnt null

                filePath = key.GetValue("Path").ToString();
                if (filePath != null)
                {
                    if (!File.Exists(filePath))
                    {
                        updatePath = true;
                        string testPath;
                        //cut up the path to get the drive replaceable
                        string[] pathParts = filePath.Split('\\');
                        //add parts back together except for the first, seperating by \
                        testPath = filePath.Remove(0, pathParts[0].Count());

                        //loop foreach drive checking if it exists
                        DriveInfo[] allDrives = DriveInfo.GetDrives();
                        //add the drive to the start and check each one
                        foreach (DriveInfo d in allDrives)
                        {
                            if (d.IsReady == true)
                            {
                                string driveName = d.Name;
                                driveName = driveName + testPath;
                                if (File.Exists(driveName))
                                { filePath = driveName; }
                            }
                        }

                        //if updating the drive doesn't find the file, then need to have the user choose a new one
                        if (!File.Exists(filePath))
                        { filePath = pickSpec(ed, doc); }
                    }
                }

                else
                {
                    //if no file existed
                    updatePath = true;
                    filePath = pickSpec(ed, doc);
                }
            }        
            #endregion

            //if the path needs to be updated, update it
            if (updatePath == true)
            { key.SetValue("Path", filePath); }

            //use a jig to insert the block
            //use path to insert the file into the DB
            BlockInserter(filePath, ed, doc);                      
        }


        public static bool regValueExist(string registryRoot, string valueName)
        {
            Microsoft.Win32.RegistryKey root = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(registryRoot, false);
            if (root.GetValue(valueName) == null)
                return false;
            else
                return true;
        }

        //loads the block into the blocktable if it isnt there
        //then uses a jig to insert at a chosen insertpoint
        public static void BlockInserter(string filePath, Editor ed, Document doc)
        {
            string specName = Path.GetFileNameWithoutExtension(filePath);
            Database dbCurrent = Application.DocumentManager.MdiActiveDocument.Database;
            using (Transaction trCurrent = dbCurrent.TransactionManager.StartTransaction())
            {
                //open block table for read
                BlockTable btCurrent = trCurrent.GetObject(dbCurrent.BlockTableId, OpenMode.ForRead) as BlockTable;

                //check if spec is already loaded into drawing
                ObjectId blkRecId = ObjectId.Null;
                if (!btCurrent.Has(specName))
                {
                    //open db to other file
                    Database db = new Database(false, true);
                    try
                    { db.ReadDwgFile(filePath, System.IO.FileShare.Read, false, ""); }
                    catch (System.Exception)
                    {
                        ed.WriteMessage("\n Unable to read drawing file.");
                        return;
                    }
                    dbCurrent.Insert(specName, db, true);

                    blkRecId = btCurrent[specName];
                }
                else
                { blkRecId = btCurrent[specName]; }

                //now insert block into current space using our jig
                Point3d stPnt = new Point3d(0, 0, 0);
                BlockReference br = new BlockReference(stPnt, blkRecId);
                BlockJig entJig = new BlockJig(br);

                //use jig
                PromptResult pr = ed.Drag(entJig);
                if(pr.Status == PromptStatus.OK)
                {
                    //add entity to the modelspace
                    BlockTableRecord btr = trCurrent.GetObject(btCurrent[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;


                    btr.AppendEntity(entJig.GetEntity());
                    trCurrent.AddNewlyCreatedDBObject(entJig.GetEntity(), true);

                    //flush graphics to see img
                    doc.TransactionManager.QueueForGraphicsFlush();
                }

                //commit changes
                trCurrent.Commit();
            }
        }

        public static string pickSpec(Editor ed, Document doc)
        {
            string spec = "";

            OpenFileDialog ofd = new OpenFileDialog(
                "Select a block to import", null,
                "dwg",
                "Pick a block", OpenFileDialog.OpenFileDialogFlags.DoNotTransferRemoteFiles);

            System.Windows.Forms.DialogResult dr = ofd.ShowDialog();
            if (dr == System.Windows.Forms.DialogResult.OK)
                spec = ofd.Filename;

            return spec;
        }
    }
}
