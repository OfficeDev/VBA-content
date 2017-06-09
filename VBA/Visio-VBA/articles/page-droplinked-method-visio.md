---
title: Page.DropLinked Method (Visio)
keywords: vis_sdr.chm10960170
f1_keywords:
- vis_sdr.chm10960170
ms.prod: visio
api_name:
- Visio.Page.DropLinked
ms.assetid: e975a150-ff48-7cae-3e3b-f21f88f2fbd2
ms.date: 06/08/2017
---


# Page.DropLinked Method (Visio)

Returns a new shape on the drawing page linked to data in a data recordset.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **DropLinked**( **_ObjectToDrop_** , **_x_** , **_y_** , **_DataRecordsetID_** , **_DataRowID_** , **_ApplyDataGraphicAfterLink_** )

 _expression_ An expression that returns a **Page** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ObjectToDrop_|Required| **[UNKNOWN]**|The object to drop. While this is typically a Visio object such as a  **Master** , **Shape** , or **Selection** object; it can be any OLE object that provides an **IDataObject** interface.|
| _x_|Required| **Double**|The x-coordinate at which to place the center of the shape's width or PinX.|
| _y_|Required| **Double**|The y-coordinate at which to place the center of the shape's height or PinY.|
| _DataRecordsetID_|Required| **Long**|The ID of the data recordset that contains the data to link to.|
| _DataRowID_|Required| **Long**|The ID of the data row that contains the data to link to.|
| _ApplyDataGraphicAfterLink_|Required| **Boolean**|Whether to apply the current data graphic to the linked shape. The default is not to apply a data graphic. See Remarks for more information.|

### Return Value

Shape


## Remarks

When you want to create shapes already linked to data on a drawing page that either does not contain any shapes or contains shapes other than the ones you want to link, you can use the  **Page.DropLinked** and **[Page.DropManyLinkedU](page-dropmanylinkedu-method-visio.md)** methods to create one or more additional shapes already linked to data. These methods resemble the existing **[Page.Drop](page-drop-method-visio.md)** and **[Page.DropManyU](page-dropmanyu-method-visio.md)** methods in that they create additional shapes at a specified location on the page; but in addition, they create links between the new shapes and specified data rows in a specified data recordset.

When the object you pass for the ObjectToDrop parameter is a shape, the center of the resulting shape's width-height box is positioned at the specified coordinates, and a  **Shape** object that represents the shape that is created is returned.

If ObjectToDrop is a  **Master** , the pin of the master is positioned at the specified coordinates. A master's pin is often, but not necessarily, at its center of rotation.

If you pass  **True** for the optional ApplyDataGraphicsAfterLink parameter, Visio applies the data graphic most recently applied to any other shape in the current document.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **DropLinked** method to create a shape on the active drawing page, centered at page coordinates (2, 2), and linked to a data row in the data recordset most recently added to the active document.

The shape passed to the  **DropLinked** method is a simple rectangle from the **Basic Shapes (US units)** stencil. Before running this macro, use the **[DataRecordsets.Add](datarecordsets-add-method-visio.md)** method or another means to add at least one data recordset to the **DataRecordsets** collection, and make sure that the **Basic Shapes (US units)** stencil is open in the Visio drawing window. In this example, the ID of the data row is set to 1; before running the code, ensure that a row with that ID exists, or change the ID value in the code.




```vb
Public Sub DropLinked_Example() 
 
    Dim vsoShape As Visio.Shape 
    Dim vsoMaster As Visio.Master 
    Dim dblX As Double 
    Dim dblY As Double  
    Dim lngDataRowID As Long 
    Dim vsoDataRecordset As Visio.DataRecordset 
    Dim intRecordesetCount As Integer 
 
    intRecordsetCount = Visio.ActiveDocument.DataRecordsets.Count 
    Set vsoDataRecordset = Visio.ActiveDocument.DataRecordsets(intRecordsetCount) 
     
    Set vsoMaster = Visio.Documents("Basic_U.VSS").Masters("Rectangle") 
    dblX = 2 
    dblY = 2 
    lngDataRowID = 1 
 
    Set vsoShape = ActivePage.DropLinked(vsoMaster, dblX, dblY, vsoDataRecordset.ID, lngDataRowID, True) 
 
End Sub
```


