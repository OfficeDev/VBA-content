---
title: Page.LinkShapesToDataRows Method (Visio)
keywords: vis_sdr.chm10960155
f1_keywords:
- vis_sdr.chm10960155
ms.prod: visio
api_name:
- Visio.Page.LinkShapesToDataRows
ms.assetid: 306c8edf-04ea-1e54-b3cf-63ea0352c242
ms.date: 06/08/2017
---


# Page.LinkShapesToDataRows Method (Visio)

Links multiple rows in the specified data recordset, as specified by their data row IDs, to multiple shapes on the page, and optionally applies the current data graphic to the linked shapes.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **LinkShapesToDataRows**( **_DataRecordsetID_** , **_DataRowIDs()_** , **_ShapeIDs()_** , **_ApplyDataGraphicAfterLink_** )

 _expression_ An expression that returns a **Page** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DataRecordsetID_|Required| **Long**|The ID of a data recordset contained in the current document containing the data to link to.|
| _DataRowIDs()_|Required| **Long**|An array of type  **Long** of data row IDs of rows in the data recordset specified in DataRecordsetID, to be linked to the shapes specified in ShapeIDs().|
| _ShapeIDs()_|Required| **Long**| An array of type **Long** of shape IDs of shapes on the page to be linked to the data rows specified in DataRowIDs() in the data recordset specified in DataRecordsetID.|
| _ApplyDataGraphicAfterLink_|Optional| **Boolean**|Whether to apply the current data graphic to the linked shapes. See Remarks for more information.|

### Return Value

Nothing


## Remarks

Index positions of the shape IDs in the array you pass for the ShapeIDs() parameter should correspond to the index position in the DataRowIDs() array of the IDs of the data rows to which you want the shapes to be linked. That is, to link the shape with ID  _1_ to the data row with ID _10_ , for example, place the shape ID and the data row ID in the same array index position in their respective arrays.

If Visio cannot establish a link between a shape and a data row, Visio skips that shape and goes on to the next shape in the array. After you run the method, to determine if all shapes in the array are actually linked to the specified data rows, call the  **[Shape.GetLinkedDataRow](shape-getlinkeddatarow-method-visio.md)** method on each shape in the array. If that method fails for any shape, it indicates that the shape is not linked to the data row. Visio will usually succeed in linking a row to a shape unless the shape is already linked to data and the link-replacement-behavior setting for the data recordset specifies that the link should not be replaced.

If you pass  **True** for the optional ApplyDataGraphicAfterLink parameter, or if you do not pass a value for this parameter, Visio re-applies the existing data graphic to shapes that already had data graphics applied before you called the method. For shapes that previously had no data graphic, Visio applies the data graphic most recently applied to any other shape in the current document.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **LinkShapesToDataRows** method to link the first three shapes added to the active drawing page to data in the first three data rows in the data recordset most recently added to the **DataRecordsets** collection of the current document. Because it does not pass a value for the optional final parameter, it also applies a data graphic to the linked shapes.

Before running this macro, open a new Visio drawing and use the  **[DataRecordsets.Add](datarecordsets-add-method-visio.md)** method or another means to add at least one data recordset to the **DataRecordsets** collection. The most recently added data recordset should contain at least three data rows. Then add at least three shapes to the drawing page.




```vb
Public Sub LinkShapesToDataRows_Example() 
 
    Dim vsoDataRecordset As Visio.DataRecordset 
    Dim intRecordsetCount As Integer 
    Dim alngDataRowIDs(0 To 2) As Long 
    Dim alngShapeIDs(0 To 2) As Long 
     
    alngShapeIDs(0) = 1 
    alngShapeIDs(1) = 2 
    alngShapeIDs(2) = 3 
     
    alngDataRowIDs(0) = 1 
    alngDataRowIDs(1) = 2 
    alngDataRowIDs(2) = 3 
         
    intRecordsetCount = Visio.ActiveDocument.DataRecordsets.Count 
    Set vsoDataRecordset = Visio.ActiveDocument.DataRecordsets(intRecordsetCount) 
     
    ActivePage.LinkShapesToDataRows vsoDataRecordset.ID, alngDataRowIDs, alngShapeIDs 
 
End Sub
```


