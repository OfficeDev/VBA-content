---
title: Shape.GetCustomPropertyLinkedColumn Method (Visio)
keywords: vis_sdr.chm11260235
f1_keywords:
- vis_sdr.chm11260235
ms.prod: visio
api_name:
- Visio.Shape.GetCustomPropertyLinkedColumn
ms.assetid: 0d6e3577-d918-1d33-135a-37a3f09f3eaa
ms.date: 06/08/2017
---


# Shape.GetCustomPropertyLinkedColumn Method (Visio)

Gets the name of the data column linked to the shape data (custom property) row in the shape's ShapeSheet spreadsheet specified by the custom property index.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **GetCustomPropertyLinkedColumn**( **_DataRecordsetID_** , **_CustomPropertyIndex_** )

 _expression_ An expression that returns a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DataRecordsetID_|Required| **Long**|The ID of the data recordset that contains the data column linked to the shape's custom property.|
| _CustomPropertyIndex_|Required| **Long**|The index of the shape data item (custom property) linked to the data column in the data recordset.|

### Return Value

String


## Remarks

If the method fails, call the  **[Shape.IsCustomPropertyLinked](shape-iscustompropertylinked-method-visio.md)** method to make sure that the shape data item (custom property row) was actually linked to the data column.


 **Note**  In some previous versions of Visio, shape data were called custom properties.


## Example

 The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the **GetCustomPropertyLinkedColumn** method to get the name of the data recordset column linked to a particular shape data item.

Before running this macro, add at least one data recordset to the  **[DataRecordsets](datarecordsets-object-visio.md)** collection of the document. The macro drops a shape onto the page, links the shape to data in the data recordset most recently added to the collection, and then tests to make sure the linking is successful. If it is, it prints the name of the data recordset column linked to the specified shape data item (custom property) in the **Immediate** window.




```vb
Public Sub GetCustomPropertyLinkedColumn_Example() 
 
    Dim vsoDataRecordset As Visio.DataRecordset 
    Dim vsoShape As Visio.Shape 
    Dim intCount As Integer 
    Dim boolIsLinked As Boolean 
    Dim lngIndex As Long 
    Dim strColumnName As String 
         
    intCount = Visio.ActiveDocument.DataRecordsets.Count 
    Set vsoDataRecordset = Visio.ActiveDocument.DataRecordsets(intCount) 
    Set vsoShape = ActivePage.DrawRectangle(2, 2, 4, 4) 
     
    vsoShape.LinkToData vsoDataRecordset.ID, 1, True 
    boolIsLinked = vsoShape.IsCustomPropertyLinked(vsoDataRecordset.ID, 1) 
     
    If boolIsLinked Then 
     
        strColumnName = vsoShape.GetCustomPropertyLinkedColumn(vsoDataRecordset.ID, 1) 
        Debug.Print "Linked column name is", strColumnName 
     
    Else 
     
        Debug.Print "Not linked." 
         
    End If 
 
End Sub
```


