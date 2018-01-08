Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
 
<CommandMethod("AddPointAndSetPointStyle")> _
Public Sub AddPointAndSetPointStyle()
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
 
      '' Create a point at (4, 3, 0) in Model space
      Dim acPoint As DBPoint = New DBPoint(New Point3d(4, 3, 0))
 
      acPoint.SetDatabaseDefaults()
 
      '' Add the new object to the block table record and the transaction
      acBlkTblRec.AppendEntity(acPoint)
      acTrans.AddNewlyCreatedDBObject(acPoint, True)
 
      '' Set the style for all point objects in the drawing
      acCurDb.Pdmode = 34
      acCurDb.Pdsize = 1
 
      '' Save the new object to the database
      acTrans.Commit()
  End Using
End Sub