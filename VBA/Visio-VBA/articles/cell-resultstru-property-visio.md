---
title: Cell.ResultStrU Property (Visio)
keywords: vis_sdr.chm10160065
f1_keywords:
- vis_sdr.chm10160065
ms.prod: visio
api_name:
- Visio.Cell.ResultStrU
ms.assetid: 2a2fc8c9-eb2c-6c49-9af6-abc120bbd610
ms.date: 06/08/2017
---


# Cell.ResultStrU Property (Visio)

Gets the value of a ShapeSheet cell expressed as a universal string. Read-only.


## Syntax

 _expression_ . **ResultStrU**( **_UnitsNameOrCode_** )

 _expression_ An expression that returns a **Cell** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _UnitsNameOrCode_|Required| **Variant**|The units to use when retrieving the value.|

### Return Value

String


## Remarks

Getting the  **ResultStrU** property is similar to getting a cell's **Result** property. The difference is that the **ResultStrU** property returns a string for the value of the cell, whereas the **Result** property returns a floating point number.

You can specify  _UnitsNameOrCode_ as an integer or a string value. If the string is invalid, an error is generated. For example, the following statements all set _UnitsNameOrCode_ to inches.

 _stringReturned_ = **Cell.ResultStrU** ( **visInches** )

 _stringReturned_ = **Cell.ResultStrU** (65)

 _stringReturned_ = **Cell.ResultStrU** ("in") where "in" can also be any of the alternate strings representing inches, such as "inch", "in.", or "intCounter".

For a complete list of valid unit strings along with corresponding Automation constants (integer values), see [About Units of Measure](http://msdn.microsoft.com/library/b6140312-b8e6-0cf2-9fe0-b14e800216bf%28Office.15%29.aspx).

Automation constants for representing units are declared by the Visio type library in member  **[VisUnitCodes ](visunitcodes-enumeration-visio.md)** .

Passing a zero (0) is sufficient for getting the value of text string cells.

You can use the  **ResultStrU** property to convert between units. For example, you can get the value in inches and then get an equivalent value in centimeters.

The  **ResultStrU** property is useful for filling controls such as edit boxes with the value of a cell.




 **Note**  Beginning with Microsoft Visio 2000, you can use both local and universal names to refer to Visio shapes, masters, documents, pages, rows, add-ons, cells, hyperlinks, styles, fonts, master shortcuts, UI objects, and layers. When a user names a shape, for example, the user is specifying a local name. Beginning with Microsoft Office Visio 2003, the ShapeSheet spreadsheet displays only universal names in cell formulas and values. (In prior versions, universal names were not visible in the user interface.)

As a developer, you can use universal names in a program when you don't want to change a name each time a solution is localized. Use the  **ResultStr** property to get an object's value expressed as a locale-specific string. Use the **ResultStrU** property to get an object's value expressed as a universal string.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows two different ways to use the  **ResultStrU** property to get the value of a ShapeSheet cell that contains a shape data item (formerly, a custom property) .



To run this macro, open a blank drawing and the  **Computers and Monitors (US Units)** stencil, and then insert a user form that contains a label, text box, and list box. Set the width of the list box to 150.




 **Note**  The  **Computers and Monitors (US Units)** stencil is available only in Visio Professional.




```vb
 
Public Sub ResultStrU_Example()  
 
    Dim vsoStencil As Visio.Document  
    Dim vsoMaster As Visio.Master  
    Dim vsoPages As Visio.Pages  
    Dim vsoPage As Visio.Page  
    Dim vsoShape As Visio.Shape  
    Dim vsoCell As Visio.Cell  
    Dim intRows As Integer 
    Dim intCounter As Integer 
 
    'Get the Pages collection for the document.  
    'ThisDocument refers to the current document.  
    Set vsoPages = ThisDocument.Pages  
 
    'Get a reference to the first page of the Pages collection.  
    Set vsoPage = vsoPages(1)  
  
    'Get the Document object for the stencil.  
    Set vsoStencil = Documents("Comps_U.VSS") 
  
    'Get the Master object for the desktop PC shape.  
    Set vsoMaster = vsoStencil.Masters("PC")  
 
    'Drop the shape in the approximate middle of the page.  
    'Coordinates passed to the Drop method are always in inches.  
    'The Drop method returns a reference to the new shape object.  
    Set vsoShape = vsoPage.Drop(vsoMaster, 4.25, 5.5)  
 
    'This example shows two methods of extracting shape data 
    'The first method retrieves the value of a shape data item by name.  
    'Note that Prop.Manufacturer implies Prop.Manufacturer.Value.  
    Set vsoCell = vsoShape.Cells("Prop.Manufacturer")  
 
    'Get the cell value as a string  
    'and put it into the text box on the form.  
    UserForm1.TextBox1.Text = vsoCell.ResultStrU(Visio.visNone)  
 
    'Set the caption of the label.  
    UserForm1.Label1.Caption = "Prop.Manufacturer"  
 
    'The second method of accessing shape data uses  
    'section, row, cell. This method is best when you want  
    'to iterate through all the shape data.  
    intRows = vsoShape.RowCount(Visio.visSectionProp)  
 
    'Make sure the list box is cleared.  
    UserForm1.ListBox1.Clear  
 
    'Loop through all the rows and add the value of Prop.Manufacturer  
    'to the list box. Rows are numbered starting with 0.  
    For intCounter = 0 To intRows - 1  
        Set vsoCell = vsoShape.CellsSRC(Visio.visSectionProp, intCounter, visCustPropsValue)  
        UserForm1.ListBox1.AddItem vsoCell.LocalName &; vbTab &; _  
            vsoCell.ResultStrU(Visio.visNone)  
    Next intCounter  
 
    'Display the user form.  
    UserForm1.Show  
 
End Sub
```


