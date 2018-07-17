---
title: Cell.ResultStr Property (Visio)
keywords: vis_sdr.chm10114230
f1_keywords:
- vis_sdr.chm10114230
ms.prod: visio
api_name:
- Visio.Cell.ResultStr
ms.assetid: f5d1236b-2596-298c-1ad4-6e19f5c32ef4
ms.date: 06/08/2017
---


# Cell.ResultStr Property (Visio)

Gets the value of a ShapeSheet cell expressed as a string. Read-only.


## Syntax

 _expression_ . **ResultStr**( **_UnitsNameOrCode_** )

 _expression_ A variable that represents a **Cell** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _UnitsNameOrCode_|Required| **Variant**|The units to use when retrieving the value.|

### Return Value

String


## Remarks

Getting the  **ResultStr** property is similar to getting a cell's **Result** property. The difference is that **ResultStr** property returns a string for the value of the cell, whereas the **Result** property returns a floating point number.

You can specify  _UnitsNameOrCode_ as an integer or a string value. If the string is invalid, an error is generated. For example, the following statements all set _UnitsNameOrCode_ to inches.

 _stringReturned_ = **Cell.ResultStr** ( **visInches** )

 _stringReturned_ = **Cell.ResultStr** (65)

 _stringReturned_ = **Cell.ResultStr** ("in") where "in" can also be any of the alternate strings representing inches, such as "inch", "in.", or "intCounter".

For a complete list of valid unit strings along with corresponding Automation constants (integer values), see [About Units of Measure](http://msdn.microsoft.com/library/b6140312-b8e6-0cf2-9fe0-b14e800216bf%28Office.15%29.aspx).

Automation constants for representing units are declared by the Visio type library in member  **[VisUnitCodes ](visunitcodes-enumeration-visio.md)** .

Passing a zero (0) is sufficient for getting the value of text string cells.

You can use the  **ResultStr** property to convert between units. For example, you can get the value in inches and then get an equivalent value in centimeters.

The  **ResultStr** property is useful for filling controls such as edit boxes with the value of a cell.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows two different ways to use the  **ResultStr** property to get the value of a ShapeSheet cell that contains a shape's shape data (formerly Custom Properties).



To run this macro, open a blank drawing and the  **Computers and Monitors (US Units)** stencil, and then insert a user form that contains a label, text box, and list box. Set the width of the list box to 150.




 **Note**  The  **Computers and Monitors (US Units)** stencil is available only in Visio Professional.




```vb
 
Public Sub ResultStr_Example() 
 
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
 
 'This example shows two methods of extracting shape data. 
 'The first method retrieves the value of a shape-data item by name. 
 'Note that Prop.Manufacturer implies Prop.Manufacturer.Value. 
 Set vsoCell = vsoShape.Cells("Prop.Manufacturer") 
 
 'Get the cell value as a string 
 'and put it into the text box on the form. 
 UserForm1.TextBox1.Text = vsoCell.ResultStr(Visio.visNone) 
 
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
 vsoCell.ResultStr(Visio.visNone) 
 Next intCounter 
 
 'Display the user form. 
 UserForm1.Show 
 
End Sub
```


