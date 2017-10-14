---
title: Shape.IsCustomPropertyLinked Method (Visio)
keywords: vis_sdr.chm11260230
f1_keywords:
- vis_sdr.chm11260230
ms.prod: visio
api_name:
- Visio.Shape.IsCustomPropertyLinked
ms.assetid: e75b910f-fb21-3e39-2ca3-ac0913adccf0
ms.date: 06/08/2017
---


# Shape.IsCustomPropertyLinked Method (Visio)

Returns whether the shape data (custom property) row in the Shape Data section of the shape's ShapeSheet spreadsheet is linked to a data row in the specified data recordset.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **IsCustomPropertyLinked**( **_DataRecordsetID_** , **_CustomPropertyIndex_** )

 _expression_ An expression that returns a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DataRecordsetID_|Required| **Long**|The ID of the data recordset that contains the data row.|
| _CustomPropertyIndex_|Required| **Long**|The index of the shape data (custom property) row in the Shape Data section of the shape's ShapeSheet.|

### Return Value

Boolean


## Remarks

Call this method before calling the  **[GetCustomPropertyLinkedColumn](shape-getcustompropertylinkedcolumn-method-visio.md)** method to make sure that the shape data item (custom property row) is actually linked to a data column.


 **Note**  In some previous versions of Visio, shape data were called custom properties.

Knowing how shapes are linked to data can help prevent conflicts and broken links when you refresh the data in one or more data recordsets.


## Example

 The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the **IsCustomPropertyLinked** method to determine whether a shape's custom property row is linked to a data row in a data recordset.

Before running this macro, add at least one data recordset to the  **[DataRecordsets](datarecordsets-object-visio.md)** collection of the document. The macro drops a shape and then tests whether the shape's first shape data item is linked to a data row in the data recordset most recently added to the collection, printing the result in the **Immediate** window. The test will fail, because the shape has not been linked to data. Then the shape is linked to data in the most recently added data recordset, and the test is run again.




```vb
Public Sub IsCustomPropertyLinked_Example() 
 
    Dim vsoDataRecordset As Visio.DataRecordset 
    Dim vsoShape As Visio.Shape 
    Dim intCount As Integer 
    Dim boolIsLinked As Boolean 
         
    intCount = Visio.ActiveDocument.DataRecordsets.Count 
    Set vsoDataRecordset = Visio.ActiveDocument.DataRecordsets(intCount) 
     
    Set vsoShape = ActivePage.DrawRectangle(2, 2, 4, 4) 
     
    boolIsLinked = vsoShape.IsCustomPropertyLinked(vsoDataRecordset.ID, 1) 
     
    Debug.Print boolIsLinked 
     
    vsoShape.LinkToData vsoDataRecordset.ID, 1, True 
    boolIsLinked = vsoShape.IsCustomPropertyLinked(vsoDataRecordset.ID, 1) 
     
    Debug.Print boolIsLinked 
     
End Sub
```


