---
title: Page.DropManyLinkedU Method (Visio)
keywords: vis_sdr.chm10960175
f1_keywords:
- vis_sdr.chm10960175
ms.prod: visio
api_name:
- Visio.Page.DropManyLinkedU
ms.assetid: 0b80591a-a563-bdad-b048-e15693410547
ms.date: 06/08/2017
---


# Page.DropManyLinkedU Method (Visio)

Creates multiple new shapes on the drawing page that are linked to multiple data rows in a data recordset. Returns the number of shape instances created and an array of IDs of those shapes.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **DropManyLinkedU**( **_ObjectsToInstance()_** , **_XYs()_** , **_DataRecordsetID_** , **_DataRowIDs()_** , **_ApplyDataGraphicAfterLink_** , **_ShapeIDs()_** )

 _expression_ An expression that returns a **Page** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ObjectsToInstance()_|Required| **Variant**|An array of type  **Variant** of objects to create instances of.|
| _XYs()_|Required| **Double**|An array of type  **Double**|
| _DataRecordsetID_|Required| **Long**|The ID of the data recordset containing the data rows to link to.|
| _DataRowIDs()_|Required| **Long**|An array of type  **Long** of IDs of the data rows containing the data to link to.|
| _ApplyDataGraphicAfterLink_|Required| **Boolean**|Whether to apply the current data graphic to the linked shapes. See Remarks for more information.|
| _ShapeIDs()_|Required| **Long**|Out parameter. An array of type  **Long** of shapes created and linked to.|

### Return Value

Long


## Remarks

When you want to create shapes already linked to data on a drawing page that either does not contain any shapes or contains shapes other than the ones you want to link, you can use the  **[Page.DropLinked](page-droplinked-method-visio.md)** and **Page.DropManyLinkedU** methods to create one or more additional shapes already linked to data. These methods resemble the existing **[Page.Drop](page-drop-method-visio.md)** and **[Page.DropManyU](page-dropmanyu-method-visio.md)** methods in that they create additional shapes at a specified location on the page; but in addition, they create links between the new shapes and specified data rows in a specified data recordset.

For the ObjectsToInstance() parameter, pass an array of objects to instance into shapes linked to data. While these objects are typically Visio objects such as  **Master** , **Shape** , or **Selection** objects, they can be any OLE objects that provide an **IDataObject** interface.

For the XYs() parameter, pass an array of type  **Double** . Each consecutive pair of array-index-position values should correspond to the _x-_ and _y-_ page coordinates where you want the instance of the object in the corresponding positon in the ObjectsToInstance() array to be positioned. For example, if you want the instance of the object in the first array index position in ObjectsToInstance() to be positioned at page coordinate (2,4), place the value _2_ in the first array index position in XYs(), and place the value _4_ in the second array index positon in that array, and so on for the rest of the objects and coordinates.

When an object you pass in the ObjectsToInstance() array is a shape, the center of the shape's width-height box is positioned at the coordinates you specify in XYs().

When an object you pass in the ObjectsToInstance() array is a master, the pin of the master is positioned at the coordinates you specify in XYs(). A master's pin is often, but not necessarily, at its center of rotation.

For the DataRowIDs() parameter, pass an array of  **Long** values that represent the IDs of the data rows in the data recordset that you want to link to the shape instances created from the objects in the corresponding array index positions in the ObjectsToInstance() array.

For the ShapeIDs() parameter, pass an empty, dimensionless array of type  **Long** . The method will return the array filled with the IDs of the newly created and linked shapes.




 **Note**   Beginning with Microsoft Visio 2000, you can use both local and universal names to refer to Visio shapes, masters, documents, pages, rows, add-ons, cells, hyperlinks, styles, fonts, master shortcuts, UI objects, and layers. When a user names a shape, for example, the user is specifying a local name.Beginning with Microsoft Office Visio 2003, the ShapeSheet spreadsheet displays only universal names in cell formulas and values. (In prior versions of Visio, universal names were not visible in the user interface.) As a developer, you can use universal names in a program when you don't want to change a name each time a solution is localized. Use the  **DropManyLinkedU** method to drop more than one shape linked to data when you are using universal names to identify the shapes.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **DropManyLinkedU** method to create several shapes on the active drawing page, centered at specified coordinates, and linked to a data rows in the data recordset most recently added to the active document. It prints the number of shapes created and their ID numbers to the **Immediate** window.

The shapes passed to the  **DropManyLinkedU** method are simple forms from the **Basic Shapes (US units)** stencil. Before running this macro, use the **[DataRecordsets.Add](datarecordsets-add-method-visio.md)** method or another means to add at least one data recordset to the **DataRecordsets** collection, and make sure that the **Basic Shapes (US units)** stencil is open in the Visio drawing window.




```vb
Sub DropManyLinkedU_Example() 
 
    Dim avarObjects(0 To 2) As Variant 
    Dim adblXYs(0 To 5) As Double   
    Dim alngDataRowIDs(0 To 2) As Long 
    Dim alngShapeIDs() As Long 
    Dim vsoDataRecordset As Visio.DataRecordset 
    Dim intRecordesetCount As Integer 
    Dim lngReturned As Long 
    Dim intCounter As Integer 
     
    intRecordsetCount = Visio.ActiveDocument.DataRecordsets.Count 
    Set vsoDataRecordset = Visio.ActiveDocument.DataRecordsets(intRecordsetCount) 
     
    Set avarObjects(0) = Visio.Documents("Basic_U.VSS").Masters("Rectangle") 
    Set avarObjects(1) = Visio.Documents("Basic_U.VSS").Masters("Triangle") 
    Set avarObjects(2) = Visio.Documents("Basic_U.VSS").Masters("Circle") 
     
    adblXYs(0) = 2 
    adblXYs(1) = 2 
    adblXYs(2) = 4 
    adblXYs(3) = 4 
    adblXYs(4) = 6 
    adblXYs(5) = 6 
         
    alngDataRowIDs(0) = 1 
    alngDataRowIDs(1) = 2 
    alngDataRowIDs(2) = 3 
         
    lngReturned = ActivePage.DropManyLinkedU(avarObjects, adblXYs, vsoDataRecordset.ID, alngDataRowIDs, True, alngShapeIDs) 
    Debug.Print lngReturned 
     
    For intCounter = 0 To lngReturned - 1 
        Debug.Print alngShapeIDs(intCounter) 
    Next 
     
End Sub
```


