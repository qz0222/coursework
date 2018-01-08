Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
 
<CommandMethod("AddLightweightPolyline")> _
Public Sub AddLightweightPolyline()
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
 
      '' Create a polyline with two segments (3 points)
      Dim acPoly As Polyline = New Polyline()
      acPoly.SetDatabaseDefaults()
      acPoly.AddVertexAt(0, New Point2d(2, 4), 0, 0, 0)
      acPoly.AddVertexAt(1, New Point2d(4, 2), 0, 0, 0)
      acPoly.AddVertexAt(2, New Point2d(6, 4), 0, 0, 0)
 
      '' Add the new object to the block table record and the transaction
      acBlkTblRec.AppendEntity(acPoly)
      acTrans.AddNewlyCreatedDBObject(acPoly, True)
 
      '' Save the new object to the database
      acTrans.Commit()
  End Using
End Sub