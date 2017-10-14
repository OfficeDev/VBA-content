---
title: Page.GetShapesLinkedToDataRow Method (Visio)
keywords: vis_sdr.chm10960150
f1_keywords:
- vis_sdr.chm10960150
ms.prod: visio
api_name:
- Visio.Page.GetShapesLinkedToDataRow
ms.assetid: d305eccc-4121-be3a-a389-f50234e526f1
ms.date: 06/08/2017
---


# Page.GetShapesLinkedToDataRow Method (Visio)

Returns an array of all shapes on the active page linked to data in the specified data row in the specified data recordset.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **GetShapesLinkedToDataRow**( **_DataRecordsetID_** , **_DataRowID_** , **_ShapeIDs()_** )

 _expression_ An expression that returns a **Page** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DataRecordsetID_|Required| **Long**|The ID of a data recordset contained in the current document.|
| _DataRowID_|Required| **Long**|The ID of a data row in the data recordset specified in DataRecordsetID.|
| _ShapeIDs()_|Required| **Long**|Out parameter. An array of type  **Long** that the method will return filled with the shape IDs of shapes on the page linked to the data row specified in DataRowID in the data recordset specified in DataRecordsetID.|

### Return Value

Nothing


## Remarks

For the ShapeIDs() parameter, pass an empty, dimensionless array of type  **Long** . If there are no shapes on the page linked to the data row specified by DataRowID in the data recordset specified by DataRecordsetID, **GetShapesLinkedToDataRow** will return an empty array.

To determine the IDs of all the data rows in a data recordset, use the  **[DataRecordset.GetDataRowIDs](datarecordset-getdatarowids-method-visio.md)** method. Note that data row IDs do not necessarily always correspond to the logical position of the data rows in the data recordset.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **GetShapesLinkedToDataRow** method to determine the shape IDs of the shapes on the page linked to data in the data row with ID number 1 in the data recordset most recently added to the **DataRecordsets** collection of the current document. It prints the shape IDs in the **Immediate** window.

Before running this macro, use the  **[DataRecordsets.Add](datarecordsets-add-method-visio.md)** method or another means to add at least one data recordset to the **DataRecordsets** collection, and make sure there is at least one shape on the active page linked to data in the data row with ID number 1 in the data recordset.




```vb
Public Sub GetShapesLinkedToDataRow_Example() 
 
    Dim vsoDataRecordset As Visio.DataRecordset 
    Dim intRecordsetCount As Integer 
    Dim alngShapeIDs() As Long 
    Dim lngDataRowID As Long 
    Dim intArrayCounter As Integer 
     
    intRecordsetCount = Visio.ActiveDocument.DataRecordsets.Count 
    Set vsoDataRecordset = Visio.ActiveDocument.DataRecordsets(intRecordsetCount) 
     
    lngDataRowID = 1 
     
    ActivePage.GetShapesLinkedToDataRow vsoDataRecordset.ID, lngDataRowID, alngShapeIDs 
     
    For intArrayCounter = LBound(alngShapeIDs) To UBound(alngShapeIDs) 
        Debug.Print alngShapeIDs(intArrayCounter) 
    Next 
     
End Sub
```


