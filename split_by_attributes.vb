Private Sub SplitByAttributes()
  ' Split Vector Layer by Attributes - ArcMap 9.2
  ' Created on 16th November 2021
  ' Author: Ben Crisford (bencrisford@gmail.com)
    
  Dim pDoc As IMxDocument
  Dim pMap As IMap
  Dim pActiveView As IActiveView
  Dim pFeatureLayer As IFeatureLayer
  Dim pFc As IFeatureClass
  
  Dim pINFeatureClassName As IFeatureClassName
  Dim pDataset As IDataset
  Dim pInDsName As IDatasetName
  Dim pFeatureSelection As IFeatureSelection
  Dim pFeatureClassName As IFeatureClassName
  Dim pOutDatasetName As IDatasetName
  Dim pWorkspaceName As IWorkspaceName
  Dim pExportOp As IExportOperation
  Dim pQueryFilter As IQueryFilter
  Dim pFeatureCursor As IFeatureCursor
  Dim pFieldName As String
  Dim pFieldIndex As Integer
      
  Set pDoc = ThisDocument
  Set pMap = pDoc.FocusMap
  Set pActiveView = pMap
  Set pFeatureLayer = pMap.Layer(0)   'Sets pFeatureLayer to be top layer in active map
  Set pFc = pFeatureLayer.FeatureClass
  
  'Get FC name from featureclass
  Set pDataset = pFc
  Set pINFeatureClassName = pDataset.FullName
  Set pInDsName = pINFeatureClassName
  pFieldName = InputBox("Enter field name:", "Split Layer by Attributes")
  pFieldIndex = pFc.fields.FindFieldByAliasName(pFieldName)
  
  'Define output feature class name, set output directory
  Set pFeatureClassName = New FeatureClassName
  Set pWorkspaceName = New WorkspaceName
  pWorkspaceName.PathName = "c:\Output"    'Define output directory
  pWorkspaceName.WorkspaceFactoryProgID = _
    "esriCore.shapefileworkspacefactory.1"
  pFeatureClassName.FeatureType = esriFTSimple
  pFeatureClassName.ShapeType = esriGeometryAny
  pFeatureClassName.ShapeFieldName = "Shape"
  
  'Get list of values
  Set pQueryFilter = New QueryFilter
  pQueryFilter.WhereClause = pFieldName & " IS NOT NULL"
  Set pFeatureCursor = pFc.Search(pQueryFilter, True)
  Dim pFeature As IFeature
  Set pFeature = pFeatureCursor.NextFeature
            
  Dim fields As New Collection, b
  Dim value As String
    
  Do Until (pFeature Is Nothing)
      value = pFeature.value(pFieldIndex)
      fields.Add value
      Set pFeature = pFeatureCursor.NextFeature
  Loop
                  
  Dim arr As New Collection, a
  Dim i As Long
  Dim count As Integer
  count = 0
   
  'Get unique values from type field
  On Error Resume Next
  For Each a In fields
     arr.Add a, a
  Next
  On Error GoTo 0
 
  For Each a In arr 
      'Query by attributes
      Set pQueryFilter = New QueryFilter
      pQueryFilter.WhereClause = pFieldName & " = '" & arr(a) & "'"
      
      'Export subset
      Set pOutDatasetName = pFeatureClassName
      pOutDatasetName.Name = arr(a)
      Set pOutDatasetName.WorkspaceName = pWorkspaceName
      Set pExportOp = New ExportOperation
      pExportOp.ExportFeatureClass pInDsName, pQueryFilter, _
            Nothing, Nothing, pOutDatasetName, 0
      count = count + 1
  Next
  
  On Error GoTo 0
  
  MsgBox ("Success. " & count & " Shapefiles exported to " & pWorkspaceName.PathName & ".")
End Sub
