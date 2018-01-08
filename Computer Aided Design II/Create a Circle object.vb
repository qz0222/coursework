Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
 
<CommandMethod("AddCircle")> _
Public Sub AddCircle()
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
 
      '' Create a circle that is at 2,3 with a radius of 4.25
      Dim acCirc As Circle = New Circle()
      acCirc.SetDatabaseDefaults()
      acCirc.Center = New Point3d(2, 3, 0)
      acCirc.Radius = 4.25
 
      '' Add the new object to the block table record and the transaction
      acBlkTblRec.AppendEntity(acCirc)
      acTrans.AddNewlyCreatedDBObject(acCirc, True)
 
      '' Save the new object to the database
      acTrans.Commit()
  End Using
End Sub