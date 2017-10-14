---
title: DataRecordset.GetMatchingRowsForRefreshConflict Method (Visio)
keywords: vis_sdr.chm16460360
f1_keywords:
- vis_sdr.chm16460360
ms.prod: visio
api_name:
- Visio.DataRecordset.GetMatchingRowsForRefreshConflict
ms.assetid: 07526278-19db-ccbc-6785-095c73128879
ms.date: 06/08/2017
---


# DataRecordset.GetMatchingRowsForRefreshConflict Method (Visio)

Returns an array of the row IDs of data-recordset rows linked to a shape that are in conflict after the data recordset is refreshed.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **GetMatchingRowsForRefreshConflict**( **_ShapeInConflict_** )

 _expression_ An expression that returns a **DataRecordset** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ShapeInConflict_|Required| **Shape**|A shape linked to data that has a conflict after the data recordset is refreshed.|

### Return Value

Long()


## Remarks

The  **[GetAllRefreshConflicts](datarecordset-getallrefreshconflicts-method-visio.md)** method returns an array of shapes where a conflict exists between data in the shape and data in the data-recordset row to which the shape is linked. To determine which data-recordset rows produced a conflict, you can then pass each of these shapes in turn to the **GetMatchingRowsForRefreshConflict** method, which returns an array of rows that are in conflict for a given shape.

Rows in the data recordset can be in conflict when two or more of them have identical primary keys, and may link to the same shape. When this occurs,  **GetMatchingRowsForRefreshConflict** returns an array containing at least two row IDs.

Conflicts can also occur when a previously data-linked row from the data recordset is removed. When this occurs, the method returns an empty array.

To remove a conflict, pass the shape that has the conflict to the  **[RemoveRefreshConflict](datarecordset-removerefreshconflict-method-visio.md)** method, which removes the conflicting information from the current document. Information about conflicts is persisted in the current document until either you delete the shape in conflict or you call **RemoveRefreshConflict** on the shape.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **GetAllRefreshConflicts** method to determine which shapes, if any, have conflicts after a data recordset is refreshed and then use the **GetMatchingRowsForRefreshConflict** method to determine which data recordset row or rows is causing the conflict. It refreshes the data recordset most recently added and if it finds no resulting conflicts, prints "No conflicts" in the **Immediate** window. If it does find conflicts, it passes each of the shapes that have conflicts to the **GetMatchingRowsForRefreshConflict** method and prints the resulting row IDs in the same window.

Before running this macro, make sure that the most recently added data recordset is a connected (non-XML-based) data recordset and that the connection to the original data source is still available. Then delete a row in the data source or make another change that will cause a conflict when you refresh the data recordset.




```vb
Public Sub GetMatchingRowsForRefreshConflict_Example() 
 
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
            
        Next intShapeCount 
         
    End If 
     
End Sub
```


