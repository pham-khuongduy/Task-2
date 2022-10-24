using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Forms;
using AcAp = Autodesk.AutoCAD.ApplicationServices.Application;


namespace test2
{
    public class AutocadAddIn
    {
        [CommandMethod("CMDPEL")]
        static public void ProjectPointToPolyline()
        {
            Document mdiActiveDocument = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            Database database = mdiActiveDocument.Database;
            Editor editor = mdiActiveDocument.Editor;

            PromptEntityOptions promptEntityOptions = new PromptEntityOptions("\nPlease select 3D polyline.");
            promptEntityOptions.SetRejectMessage("\nPlease select 3D polyline.");
            promptEntityOptions.AddAllowedClass(typeof(Polyline3d), true);
            PromptEntityResult entity = editor.GetEntity(promptEntityOptions);
            if (entity.Status == PromptStatus.Error) return;

            using (Transaction transaction = database.TransactionManager.StartTransaction())
            {
                Polyline3d polyline3D = transaction.GetObject(entity.ObjectId, OpenMode.ForWrite) as Polyline3d;
                PromptStringOptions promptStringOptions = new PromptStringOptions("\nPlease enter the length");
                promptStringOptions.AllowSpaces = true;
                PromptResult promptResult = editor.GetString(promptStringOptions);
                var checkLength = double.TryParse(promptResult.StringResult, out double result);
                if (!checkLength)
                {
                    MessageBox.Show("Input value was not correct. Please try again.");
                    return;
                }
                if(result> polyline3D.Length)
                {
                    MessageBox.Show("Input value is out of 3DPolyline. Please try again.");
                    return;
                }
                if (result < 0)
                {
                    MessageBox.Show("Input value was not correct. Please try again.");
                    return;
                }
                if (polyline3D != null)
                {
                    var poly2D = polyline3D.GetProjectedCurve(new Plane(Point3d.Origin, Vector3d.ZAxis), Vector3d.ZAxis);
                    var pointAtDist = poly2D.GetPointAtDist(result);
                    var pointOnPoly = polyline3D.GetClosestPointTo(pointAtDist, Vector3d.ZAxis, true);


                    MessageBox.Show("The elevation at lenght is: " + Math.Round(pointOnPoly.Z, 3));
                }


                transaction.Commit();
            }
        }
    }   
}
