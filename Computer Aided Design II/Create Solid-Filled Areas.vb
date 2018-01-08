Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
 
<CommandMethod("Add2DSolid")> _
Public Sub Add2DSolid()
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
 
      '' Create a quadrilateral (bow-tie) solid in Model space
      Dim ac2DSolidBow As Solid = New Solid(New Point3d(0, 0, 0), _
                                            New Point3d(5, 0, 0), _
                                            New Point3d(5, 8, 0), _
                                            New Point3d(0, 8, 0))
 
      ac2DSolidBow.SetDatabaseDefaults()
 
      '' Add the new object to the block table record and the transaction
      acBlkTblRec.AppendEntity(ac2DSolidBow)
      acTrans.AddNewlyCreatedDBObject(ac2DSolidBow, True)
 
      '' Create a quadrilateral (square) solid in Model space
      Dim ac2DSolidSqr As Solid = New Solid(New Point3d(10, 0, 0), _
                                            New Point3d(15, 0, 0), _
                                            New Point3d(10, 8, 0), _
                                            New Point3d(15, 8, 0))
 
      ac2DSolidSqr.SetDatabaseDefaults()
 
      '' Add the new object to the block table record and the transaction
      acBlkTblRec.AppendEntity(ac2DSolidSqr)
      acTrans.AddNewlyCreatedDBObject(ac2DSolidSqr, True)
 
      '' Save the new object to the database
      acTrans.Commit()
  End Using
End Sub