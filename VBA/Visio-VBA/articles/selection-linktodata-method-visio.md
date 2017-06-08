---
title: Selection.LinkToData Method (Visio)
keywords: vis_sdr.chm11160190
f1_keywords:
- vis_sdr.chm11160190
ms.prod: visio
api_name:
- Visio.Selection.LinkToData
ms.assetid: 1aa42548-2f3a-015d-e618-c0e103ffaea3
ms.date: 06/08/2017
---


# Selection.LinkToData Method (Visio)

Links a selection of shapes to a single data row in a data recordset.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **LinkToData**( **_DataRecordsetID_** , **_DataRowID_** , **_AutoApplyDataGraphics_** )

 _expression_ An expression that returns a **Selection** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DataRecordsetID_|Required| **Long**|The ID of the data recordset containing the data to link to.|
| _DataRowID_|Required| **Long**|The ID of the row in the data recordset containing the particular data record to link to. |
| _AutoApplyDataGraphics_|Optional| **Boolean**|Whether to automatically apply a data graphic to the linked shapes. See Remarks for more information.|

### Return Value

Nothing


## Remarks

The  **Selection.LinkToData** method functions much like the same method of the **Shape** object, **[Shape.LinkToData](shape-linktodata-method-visio.md)** , except that it links a selection of shapes, instead of a single shape, to a single data row.

If Visio cannot establish a link between a shape and the data row, Visio skips that shape and goes on to the next shape in the selection. After you run the method, to determine if all shapes in the selection are actually linked to the data row, call the  **[Shape.GetLinkedDataRow](shape-getlinkeddatarow-method-visio.md)** method on each shape in the selection. If that method fails for any shape, it indicates that the shape is not linked to the data row. Visio will usually succeed in linking a row to a shape unless the shape is already linked to data and the link-replacement-behavior setting for the data recordset specifies that the link should not be replaced.

If you pass  **True** for the AutoApplyDataGraphics parameter, Visio re-applies the existing data graphic to shapes that already had data graphics applied before you called the method. For shapes that previously had no data graphic, Visio applies the data graphic most recently applied to any other shape in the current document.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **LinkToData** method to link the shapes in a selection to a data row.

Before running this macro, place several shapes on the page and add at least one data recordset to the  **[DataRecordsets](datarecordsets-object-visio.md)** collection of the document. The macro uses the ID of the data recordset most recently added to the collection. It links selected shapes to the data in the first row of the data recordset.




```vb
Public Sub LinkToData_Example() 
 
    Dim vsoDataRecordset As Visio.DataRecordset 
    Dim vsoSelection As Visio.Selection 
    Dim intCount As Integer 
     
    intCount = Visio.ActiveDocument.DataRecordsets.Count 
    Set vsoDataRecordset = Visio.ActiveDocument.DataRecordsets(intCount) 
     
    ActiveWindow.DeselectAll 
    ActiveWindow.SelectAll 
     
    Set vsoSelection = ActiveWindow.Selection 
    vsoSelection.LinkToData vsoDataRecordset.ID, 1, True 
 
End Sub
```


