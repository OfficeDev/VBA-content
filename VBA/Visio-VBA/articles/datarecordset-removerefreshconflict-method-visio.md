---
title: DataRecordset.RemoveRefreshConflict Method (Visio)
keywords: vis_sdr.chm16460355
f1_keywords:
- vis_sdr.chm16460355
ms.prod: visio
api_name:
- Visio.DataRecordset.RemoveRefreshConflict
ms.assetid: a92abdb7-f47c-b843-cacf-6acca68d9c66
ms.date: 06/08/2017
---


# DataRecordset.RemoveRefreshConflict Method (Visio)

Clears information about a conflict for a data-linked shape from the current document.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **RemoveRefreshConflict**( **_ShapeInConflict_** )

 _expression_ An expression that returns a **DataRecordset** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ShapeInConflict_|Required| **[IVSHAPE]**|A shape linked to data that has a conflict after the data recordset is refreshed.|

### Return Value

Nothing


## Remarks

If ShapeInConflict actually has no conflicts, the method will have no effect.

If you choose not to remove information about a conflict, that information will be persisted in the current document indefinitely.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **RemoveRefreshConflict** method to remove conflicts. It calls the **GetAllRefreshConflicts** method to determine which shapes, if any, have conflicts after a data recordset is refreshed. Then it calls the **GetMatchingRowsForRefreshConflict** method to determine which data recordset row or rows is causing the conflict.

It refreshes the data recordset most recently added and if it finds no resulting conflicts, prints "No conflicts" in the  **Immediate** window. If it does find conflicts, it passes each of the shapes that have conflicts to the **GetMatchingRowsForRefreshConflict** method and prints the resulting row IDs in the same window. Then it passes the shapes that have conflicts to the **RemoveRefreshConflict** method to remove the conflicts.

Before running this macro, make sure that the most recently added data recordset is a connected (non-XML-based) data recordset and that the connection to the original data source is still available. Then delete a row in the data source or make another change that will cause a conflict when you refresh the data recordset.




```vb
Public Sub RemoveRefreshConflicts_Example() 
 
    Dim vsoDataRecordset As Visio.DataRecordset 
    Dim intRecordsetCount As Integer 
    Dim intShapeCount As Integer 
    Dim vsoShapes() As Visio.Shape 
     
    Dim intRowCount As Integer 
    Dim vsoShapeInConflict As Visio.Shape 
     
    Dim alngRowIDs() As Long 
    Dim lngvsoRowID As Long 
         
    intRecordsetCount = Visio.ActiveDocument.DataRecordsets.Count 
    Set vsoDataRecordset = Visio.ActiveDocument.DataRecordsets(intRecordsetCount) 
     
    vsoDataRecordset.Refresh 
    vsoShapes = vsoDataRecordset.GetAllRefreshConflicts 
     
    If IsEmpty(vsoShapes) Then 
        Debug.Print "No conflict" 
    Else 
        For intShapeCount = LBound(vsoShapes) To UBound(vsoShapes) 
            Set vsoShapeInConflict = vsoShapes(intShapeCount) 
            alngRowIDs = vsoDataRecordset.GetMatchingRowsForRefreshConflict(vsoShapeInConflict) 
             
            If IsEmpty(alngRowIDs) Then 
                Debug.Print "For shape:", vsoShapeInConflict.Name, "Row deleted." 
            Else 
                For intRowCount = LBound(alngRowIDs) To UBound(alngRowIDs) 
                    lngvsoRowID = alngRowIDs(intRowCount) 
                    Debug.Print "For shape:", vsoShapeInConflict.Name, "Row ID of row in conflict:", lngvsoRowID 
                Next intRowCount 
            End If 
             
            Call vsoDataRecordset.RemoveRefreshConflict (vsoShapeInConflict) 
            
        Next intShapeCount 
         
    End If 
     
End Sub
```


