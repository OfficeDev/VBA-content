---
title: DataRecordset.GetAllRefreshConflicts Method (Visio)
keywords: vis_sdr.chm16460350
f1_keywords:
- vis_sdr.chm16460350
ms.prod: visio
api_name:
- Visio.DataRecordset.GetAllRefreshConflicts
ms.assetid: 96d1c866-6c0d-f750-46a8-8257340ebd71
ms.date: 06/08/2017
---


# DataRecordset.GetAllRefreshConflicts Method (Visio)

Returns an array that contains shapes linked to data rows that have non-resolved conflicts after a data recordset is refreshed. .


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **GetAllRefreshConflicts**

 _expression_ An expression that returns a **DataRecordset** object.


### Return Value

Shape()


## Remarks

When you refresh data from a data source that has changed since the last time you refreshed data, conflicts can occur. Conflicts can result when a single shape is linked to more than one row in the same data source, or when a shape is linked to a row in the data source that has been deleted. 

To determine which data-recordset rows produced the conflict, pass the shape(s) returned by  **GetAllRefreshConflicts** to the **[GetMatchingRowsForRefreshConflict](datarecordset-getmatchingrowsforrefreshconflict-method-visio.md)** method, which returns an array of rows that are in conflict.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **GetAllRefreshConflicts** method to determine which shapes, if any, have conflicts after a data recordset is refreshed. It refreshes the data recordset most recently added and if it finds no resulting conflicts, prints "No conflicts" in the **Immediate** window. If it does find conflicts, it iterates through the returned shape array and prints the names of the shapes that have conflicts in the same window.

Before running this macro, make sure that the most recently added data recordset is a connected (non-XML-based) data recordset and that the connection to the original data source is still available. Then delete a row from the data source or make another change that will cause a conflict.




```vb
Public Sub GetAllRefreshConflicts_Example() 
 
    Dim vsoDataRecordset As Visio.DataRecordset 
    Dim intRecordsetCount As Integer 
    Dim intShapeCount As Integer         
    Dim avsoShapes() As Visio.Shape     
     
    intRecordsetCount = Visio.ActiveDocument.DataRecordsets.Count 
    Set vsoDataRecordset = Visio.ActiveDocument.DataRecordsets(intRecordsetCount) 
 
    vsoDataRecordset.Refresh    
    avsoShapes = vsoDataRecordset.GetAllRefreshConflicts 
     
    If IsEmpty(avsoShapes) Then 
        Debug.Print "No conflict" 
    Else 
        For intShapeCount = LBound(avsoShapes) To UBound(avsoShapes) 
            Debug.Print avsoShapes(intShapeCount).Name 
        Next intShapeCount 
    End If 
     
End Sub
```


