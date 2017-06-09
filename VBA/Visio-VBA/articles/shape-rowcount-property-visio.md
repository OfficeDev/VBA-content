---
title: Shape.RowCount Property (Visio)
keywords: vis_sdr.chm11214245
f1_keywords:
- vis_sdr.chm11214245
ms.prod: visio
api_name:
- Visio.Shape.RowCount
ms.assetid: 358f07c8-f72a-134a-53d8-9b70f2400484
ms.date: 06/08/2017
---


# Shape.RowCount Property (Visio)

Returns the number of rows in a ShapeSheet section. Read-only.


## Syntax

 _expression_ . **RowCount**( **_Section_** )

 _expression_ A variable that represents a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Section_|Required| **Integer**|The section whose rows to count.|

### Return Value

Integer


## Remarks

The  _Section_ argument must be a section constant. For a list of section constants, see the **AddSection** method.

Use the  **RowCount** property primarily with sections that contain a variable number of rows, such as Geometry and Connection Points sections. For sections that have a fixed number of rows, the **RowCount** property returns the number of rows in the section that possess at least one cell whose value is local to the shape, as opposed to rows whose cells are all inherited from a master or style. Inheriting from a master or style is usually better because Microsoft Office Visio does not need to store as much information. In the ShapeSheet window, cells with local values appear in blue, and cells with inherited values appear in black. You can use the **IsInherited** property to determine if a cell is inherited.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **RowCount** property to find the number of ShapeSheet rows to iterate through.



To run this macro, open a blank drawing and the  **Computers and Monitors (US Units)** stencil, and then insert a user form that contains a label, text box, and list box. Set the width of the list box to 150.


 **Note**  The  **Computers and Monitors (US Units)** stencil is available only in Microsoft Office Visio Professional.




```vb
 
Public Sub RowCount_Example() 
 
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
 Set vsoStencil = Documents("COMPS_U.VSS") 
 
 'Get the Master object for the desktop PC shape. 
 Set vsoMaster = vsoStencil.Masters("PC") 
 
 'Drop the shape in the approximate middle of the page. 
 'Coordinates passed to the Drop method are always in inches. 
 'The Drop method returns a reference to the new shape object. 
 Set vsoShape = vsoPage.Drop(vsoMaster, 4.25, 5.5) 
 
 'This example shows two methods of extracting custom 
 'properties. The first method retrieves the value of a custom 
 'property by name. 
 'Note that Prop.Manufacturer implies Prop.Manufacturer.Value. 
 Set vsoCell = vsoShape.Cells("Prop.Manufacturer") 
 
 'Get the cell value as a string 
 'and put it into the text box on the form. 
 UserForm1.TextBox1.Text = vsoCell.ResultStr(Visio.visNone) 
 
 'Set the caption of the label. 
 UserForm1.Label1.Caption = "Prop.Manufacturer" 
 
 'The second method of accessing custom properties uses 
 'section, row, cell. This method is best when you want 
 'to iterate through all the properties. 
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


