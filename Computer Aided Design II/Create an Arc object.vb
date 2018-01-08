Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
 
<CommandMethod("AddArc")> _
Public Sub AddArc()
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
 
      '' Create an arc that is at 6.25,9.125 with a radius of 6, and
      '' starts at 64 degrees and ends at 204 degrees
      Dim acArc As Arc = New Arc(New Point3d(6.25, 9.125, 0), _
                                 6, 1.117, 3.5605)
 
      acArc.SetDatabaseDefaults()
 
      '' Add the new object to the block table record and the transaction
      acBlkTblRec.AppendEntity(acArc)
      acTrans.AddNewlyCreatedDBObject(acArc, True)
 
      '' Save the new object to the database
      acTrans.Commit()
  End Using
End Sub