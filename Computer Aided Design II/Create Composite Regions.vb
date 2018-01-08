Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
 
<CommandMethod("CreateCompositeRegions")> _
Public Sub CreateCompositeRegions()
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
 
      '' Create two in memory circles
      Dim acCirc1 As Circle = New Circle()
      acCirc1.SetDatabaseDefaults()
      acCirc1.Center = New Point3d(4, 4, 0)
      acCirc1.Radius = 2
 
      Dim acCirc2 As Circle = New Circle()
      acCirc2.SetDatabaseDefaults()
      acCirc2.Center = New Point3d(4, 4, 0)
      acCirc2.Radius = 1
 
      '' Adds the circle to an object array
      Dim acDBObjColl As DBObjectCollection = New DBObjectCollection()
      acDBObjColl.Add(acCirc1)
      acDBObjColl.Add(acCirc2)
 
      '' Calculate the regions based on each closed loop
      Dim myRegionColl As DBObjectCollection = New DBObjectCollection()
      myRegionColl = Region.CreateFromCurves(acDBObjColl)
      Dim acRegion1 As Region = myRegionColl(0)
      Dim acRegion2 As Region = myRegionColl(1)
 
      '' Subtract region 1 from region 2
      If acRegion1.Area > acRegion2.Area Then
          '' Subtract the smaller region from the larger one
          acRegion1.BooleanOperation(BooleanOperationType.BoolSubtract, acRegion2)
          acRegion2.Dispose()
 
          '' Add the final region to the database
          acBlkTblRec.AppendEntity(acRegion1)
          acTrans.AddNewlyCreatedDBObject(acRegion1, True)
      Else
          '' Subtract the smaller region from the larger one
          acRegion2.BooleanOperation(BooleanOperationType.BoolSubtract, acRegion1)
          acRegion1.Dispose()
 
          '' Add the final region to the database
          acBlkTblRec.AppendEntity(acRegion2)
          acTrans.AddNewlyCreatedDBObject(acRegion2, True)
      End If
 
      '' Dispose of the in memory objects not appended to the database
      acCirc1.Dispose()
      acCirc2.Dispose()
 
      '' Save the new object to the database
      acTrans.Commit()
  End Using
End Sub