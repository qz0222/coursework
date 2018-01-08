Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
 
<CommandMethod("AddRegion")> _
Public Sub AddRegion()
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
 
      '' Create an in memory circle
      Using acCirc As Circle = New Circle()
          acCirc.SetDatabaseDefaults()
          acCirc.Center = New Point3d(2, 2, 0)
          acCirc.Radius = 5
 
          '' Adds the circle to an object array
          Dim acDBObjColl As DBObjectCollection = New DBObjectCollection()
          acDBObjColl.Add(acCirc)
 
          '' Calculate the regions based on each closed loop
          Dim myRegionColl As DBObjectCollection = New DBObjectCollection()
          myRegionColl = Region.CreateFromCurves(acDBObjColl)
          Dim acRegion As Region = myRegionColl(0)
 
          '' Add the new object to the block table record and the transaction
          acBlkTblRec.AppendEntity(acRegion)
          acTrans.AddNewlyCreatedDBObject(acRegion, True)
 
          '' Dispose of the in memory object not appended to the database
      End Using
 
      '' Save the new object to the database
      acTrans.Commit()
  End Using
End Sub