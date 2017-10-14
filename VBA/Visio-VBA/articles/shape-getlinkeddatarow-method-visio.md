---
title: Shape.GetLinkedDataRow Method (Visio)
keywords: vis_sdr.chm11260215
f1_keywords:
- vis_sdr.chm11260215
ms.prod: visio
api_name:
- Visio.Shape.GetLinkedDataRow
ms.assetid: 55e578a5-da95-9a5c-3d1d-5cc5edeb57a7
ms.date: 06/08/2017
---


# Shape.GetLinkedDataRow Method (Visio)

Gets the ID of the data row in the specified data recordset linked to the shape.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **GetLinkedDataRow**( **_DataRecordsetID_** )

 _expression_ An expression that returns a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DataRecordsetID_|Required| **Long**|The ID of the data recordset that contains the linked row.|

### Return Value

Long


## Remarks

The  **GetLinkedDataRow** method fails if the shape is not linked to a data row.


## Example

 The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the **GetLinkedDataRow** method to get the ID of the data row in the specified data recordset linked to the shape.

Before running this macro, add at least one data recordset to the  **[DataRecordsets](datarecordsets-object-visio.md)** collection of the document. The macro drops a shape onto the page, links the shape to a data row in the data recordset most recently added to the collection, gets the ID of the row, and then prints the ID of the row in the **Immediate** window.




```vb
Public Sub GetLinkedDataRow_Example() 
 
    Dim vsoDataRecordset As Visio.DataRecordset 
    Dim vsoShape As Visio.Shape 
    Dim intCount As Integer 
    Dim lngRowID As Long 
     
    intCount = Visio.ActiveDocument.DataRecordsets.Count 
    Set vsoDataRecordset = Visio.ActiveDocument.DataRecordsets(intCount) 
     
    Set vsoShape = ActivePage.DrawRectangle(2, 2, 4, 4) 
         
    vsoShape.LinkToData vsoDataRecordset.ID, 1, True 
            
    lngRowID = vsoShape.GetLinkedDataRow(vsoDataRecordset.ID) 
    Debug.Print lngRowID 
         
End Sub
```


