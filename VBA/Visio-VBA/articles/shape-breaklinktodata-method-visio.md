---
title: Shape.BreakLinkToData Method (Visio)
keywords: vis_sdr.chm11260195
f1_keywords:
- vis_sdr.chm11260195
ms.prod: visio
api_name:
- Visio.Shape.BreakLinkToData
ms.assetid: 1f4ed559-061e-f016-739c-e760e634dba8
ms.date: 06/08/2017
---


# Shape.BreakLinkToData Method (Visio)

Breaks the link between the shape and the data row to which it is linked in the specified data recordset.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **BreakLinkToData**( **_DataRecordsetID_** )

 _expression_ An expression that returns a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DataRecordsetID_|Required| **Long**|The ID of the data recordset containing the data row the shape is linked to.|

### Return Value

Nothing


## Remarks

Breaking the link between an shape and a data row does not remove shape data (called custom properties in some previous versions of Visio) from the shape, nor does it remove any data graphics associated with the shape.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **BreakLinkToData** method to break the link between a shape and a data row in a data recordset.

Before running this macro, place a shape on the page, add at least one data recordset to the  **[DataRecordsets](datarecordsets-object-visio.md)** collection of the document, and use any linking method to link the shape to a data row in the data recordset you most recently added to the collection. Alternatively, you could link the shape to a data row by dragging the row from the **External Data** window onto the shape in the Visio user interface. Then select the linked shape.




```vb
Public Sub BreakLinkToData_Example() 
 
    Dim vsoDataRecordset As Visio.DataRecordset 
    Dim vsoShape As Visio.Shape 
    Dim intCount As Integer 
         
    intCount = Visio.ActiveDocument.DataRecordsets.Count 
    Set vsoDataRecordset = Visio.ActiveDocument.DataRecordsets(intCount) 
     
    Set vsoShape = ActiveWindow.Selection.PrimaryItem 
    vsoShape.BreakLinkToData (vsoDataRecordset.ID) 
    
End Sub
```


