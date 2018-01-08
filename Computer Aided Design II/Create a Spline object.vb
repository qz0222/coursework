Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
 
<CommandMethod("AddSpline")> _
Public Sub AddSpline()
  '' Get the current document and database
  Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
  Dim acCurDb As Database = acDoc.Database
 
  '' Start a transaction
  Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
 
      '' Open the Block table for read
      Dim acBlkTbl As BlockTable
      acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
 
      '' Open the Block table record Model space for write
      Dim acBlkTblRec As BlockTableRecord
      acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), _
                                      OpenMode.ForWrite)
 
      '' Define the fit points for the spline
      Dim ptColl As Point3dCollection = New Point3dCollection()
      ptColl.Add(New Point3d(0, 0, 0))
      ptColl.Add(New Point3d(5, 5, 0))
      ptColl.Add(New Point3d(10, 0, 0))
 
      '' Get a 3D vector from the point (0.5,0.5,0)
      Dim vecTan As Vector3d = New Point3d(0.5, 0.5, 0).GetAsVector
 
      '' Create a spline through (0, 0, 0), (5, 5, 0), and (10, 0, 0) with a
      '' start and end tangency of (0.5, 0.5, 0.0)
      Dim acSpline As Spline = New Spline(ptColl, vecTan, vecTan, 4, 0.0)
 
      acSpline.SetDatabaseDefaults()
 
      '' Add the new object to the block table record and the transaction
      acBlkTblRec.AppendEntity(acSpline)
      acTrans.AddNewlyCreatedDBObject(acSpline, True)
 
      '' Save the new object to the database
      acTrans.Commit()
  End Using
End Sub