---
title: Selection.BreakLinkToData Method (Visio)
keywords: vis_sdr.chm11160195
f1_keywords:
- vis_sdr.chm11160195
ms.prod: visio
api_name:
- Visio.Selection.BreakLinkToData
ms.assetid: 83a52ed7-1d10-9005-4a1a-339995106d8b
ms.date: 06/08/2017
---


# Selection.BreakLinkToData Method (Visio)

Breaks links between all shapes in the selection and data rows in the specified data recordset.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **BreakLinkToData**( **_DataRecordsetID_** )

 _expression_ An expression that returns a **Selection** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DataRecordsetID_|Required| **Long**|The ID of the data recordset containing the data rows with which to break links.|

### Return Value

Nothing


## Remarks

If Microsoft Visio cannot break the link between a shape in the selection and the data row, or if the link does not exist, Visio skips that shape and goes on to the next shape in the selection. After you run the method, to determine if anyshapes in the selection are still linked to a data row, call the  **[Shape.GetLinkedDataRow](shape-getlinkeddatarow-method-visio.md)** method on each shape in the selection. If the **GetLinkedDataRow** method fails for any shape, it indicates that the shape either no longer is linked to the data row, or never was linked to the data row.

Note that breaking links between shapes and data does not remove shape data (called custom properties in some previous versions of Visio) from shapes, nor does it remove data graphics associated with shapes.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **BreakLinkToData** method to break all links between the shapes in a selection and data rows in a data recordset.

Before running this macro, place several shapes on the page, add at least one data recordset to the  **[DataRecordsets](datarecordsets-object-visio.md)** collection of the document, and use any linking method to link several shapes to one or more data rows in the data recordset you most recently added to the collection.




```vb
Public Sub BreakLinkToData_Example() 
 
    Dim vsoDataRecordset As Visio.DataRecordset 
    Dim vsoSelection As Visio.Selection 
    Dim intCount As Integer 
        
    intCount = Visio.ActiveDocument.DataRecordsets.Count 
    Set vsoDataRecordset = Visio.ActiveDocument.DataRecordsets(intCount) 
     
    ActiveWindow.DeselectAll 
    ActiveWindow.SelectAll 
     
    Set vsoSelection = ActiveWindow.Selection 
    Call vsoSelection.BreakLinkToData(vsoDataRecordset.ID) 
     
End Sub
```


