---
title: Shape.LinkToData Method (Visio)
keywords: vis_sdr.chm11260190
f1_keywords:
- vis_sdr.chm11260190
ms.prod: visio
api_name:
- Visio.Shape.LinkToData
ms.assetid: 75dd1543-e643-0c7d-a89a-f0dd09d6d323
ms.date: 06/08/2017
---


# Shape.LinkToData Method (Visio)

Links a shape to a data row in a data recordset.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **LinkToData**( **_DataRecordsetID_** , **_RowID_** , **_AutoApplyDataGraphics_** )

 _expression_ An expression that returns a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DataRecordsetID_|Required| **Long**|The ID of the data recordset containing the data to link to.|
| _RowID_|Required| **Long**|The ID of the row in the data recordset containing the particular data record to link to. |
| _AutoApplyDataGraphics_|Optional| **Boolean**|Whether to automatically apply a data graphic to the linked shapes. See Remarks for more information.|

### Return Value

Nothing


## Remarks

The  **Shape.LinkToData** method functions much like the same method of the **Selection** object, **[Selection.LinkToData](selection-linktodata-method-visio.md)** , except that it links a single shape, instead of a selection of shapes, to a single data row.

If you pass  **True** for the AutoApplyDataGraphics parameter, Microsoft Visio re-applies the existing data graphic to a shape if it already had a data graphic applied before you called the method. For a shape that previously had no data graphic, Visio applies the data graphic most recently applied to any other shape in the current document.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **LinkToData** method to link a shape to a data row.

Before running this macro, add at least one data recordset to the  **[DataRecordsets](datarecordsets-object-visio.md)** collection of the document. The macro uses the ID of the data recordset most recently added to the collection. It draws a rectangle shape on the page and links it to the data in the first row of the data recordset, while applying the default data graphic to the shape.




```vb
Public Sub LinkToData_Example() 
 
    Dim vsoDataRecordset As Visio.DataRecordset 
    Dim vsoShape As Visio.Shape 
    Dim intCount As Integer 
     
    intCount = Visio.ActiveDocument.DataRecordsets.Count 
    Set vsoDataRecordset = Visio.ActiveDocument.DataRecordsets(intCount) 
     
    Set vsoShape = ActivePage.DrawRectangle(2, 2, 5, 5) 
     
    vsoShape.LinkToData vsoDataRecordset.ID, 1, True 
 
End Sub
```


