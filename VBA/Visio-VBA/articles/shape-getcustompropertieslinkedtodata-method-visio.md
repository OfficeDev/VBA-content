---
title: Shape.GetCustomPropertiesLinkedToData Method (Visio)
keywords: vis_sdr.chm11260225
f1_keywords:
- vis_sdr.chm11260225
ms.prod: visio
api_name:
- Visio.Shape.GetCustomPropertiesLinkedToData
ms.assetid: 8a0d783d-f5ee-d6c0-adbd-377cbe65e5f5
ms.date: 06/08/2017
---


# Shape.GetCustomPropertiesLinkedToData Method (Visio)

Gets the IDs of the shape-data-item (custom property) rows in the Shape Data section of the shape's ShapeSheet spreadsheet linked to the specified data recordset.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **GetCustomPropertiesLinkedToData**( **_DataRecordsetID_** , **_CustomPropertyIndices()_** )

 _expression_ An expression that returns a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DataRecordsetID_|Required| **Long**|The ID of the data recordset that contains the data the shape is linked to.|
| _CustomPropertyIndices()_|Required| **Long**|Out parameter. An empty, dimensionless array that the method fills with the row IDs of the shape-data-item (custom property) rows in the shape's ShapeSheet that are linked to data columns in the data recordset.|

### Return Value

Nothing


## Remarks

Knowing how shapes are linked to data can help prevent conflicts and broken links when you refresh the data in one or more data recordsets.


 **Note**  In some previous versions of Visio, shape data were called custom properties.


## Example

 The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the **GetCustomPropertiesLinkedToData** method to get the IDs of the shape-data-item (custom property) rows linked to a data column in a data recordset.

Before running this macro, add at least one data recordset to the  **[DataRecordsets](datarecordsets-object-visio.md)** collection of the document. The macro drops a shape onto the page, links the shape to data in the data recordset most recently added to the collection, and then tests to make sure the linking is successful. If it is, it gets the row IDs of all ShapeSheet rows linked to data prints the IDs of the rows in the **Immediate** window.




```vb
Public Sub GetCustomPropertiesLinkedToData_Example() 
 
    Dim vsoDataRecordset As Visio.DataRecordset 
    Dim vsoShape As Visio.Shape 
    Dim intCount As Integer 
    Dim boolIsLinked As Boolean 
    Dim alngIndices() As Long 
    Dim intArrayIndex as Integer 
            
    intCount = Visio.ActiveDocument.DataRecordsets.Count 
    Set vsoDataRecordset = Visio.ActiveDocument.DataRecordsets(intCount) 
    Set vsoShape = ActivePage.DrawRectangle(2, 2, 4, 4) 
         
    vsoShape.LinkToData vsoDataRecordset.ID, 1, True 
    boolIsLinked = vsoShape.IsCustomPropertyLinked(vsoDataRecordset.ID, 1) 
     
    If boolIsLinked Then 
         
        vsoShape.GetCustomPropertiesLinkedToData vsoDataRecordset.ID, alngIndices 
        For intArrayIndex = LBound(alngIndices) To UBound(alngIndices) 
             Debug.Print alngIndices(intArrayIndex) 
        Next 
     
    Else 
     
        Debug.Print "Not linked." 
         
    End If 
 
End Sub
```


