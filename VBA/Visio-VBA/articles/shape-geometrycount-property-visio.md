---
title: Shape.GeometryCount Property (Visio)
keywords: vis_sdr.chm11213600
f1_keywords:
- vis_sdr.chm11213600
ms.prod: visio
api_name:
- Visio.Shape.GeometryCount
ms.assetid: 4dffe649-3629-6e3e-bcc0-d860eb1efdbe
ms.date: 06/08/2017
---


# Shape.GeometryCount Property (Visio)

Returns the number of Geometry sections for a shape. Read-only.


## Syntax

 _expression_ . **GeometryCount**

 _expression_ A variable that represents a **Shape** object.


### Return Value

Integer


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **GeometryCount** property to determine the number of Geometry sections a shape has.



To run this macro, first insert into your project a user form that contains a list box. Use the default names for the form and the list box. In the  **Properties** window, set the width of the form to 400 and the width of the list box to 300. This macro also assumes that you have one or more shapes on the active page.




```vb
Public Sub GeometryCount_Example() 
 Dim vsoShape As Visio.Shape 
 Dim intCurrentGeometrySection As Integer 
 Dim intCurrentGeometrySectionIndex As Integer 
 Dim intRows As Integer 
 Dim intCells As Integer 
 Dim intCurrentRow As Integer 
 Dim intCurrentCell As Integer 
 Dim intSections As Integer 
 
 'Get the first shape from the active page. 
 Set vsoShape = ActivePage.Shapes(1) 
 
 'Clear the list box. 
 UserForm1.ListBox1.Clear 
 
 'Get the count of Geometry sections in the shape. 
 '(If the shape is a group, this will be 0.) 
 intSections = vsoShape.GeometryCount 
 
 'Iterate through all Geometry sections for the shape. 
 'Because we are adding the current Geometry section index to 
 'the constant visSectionFirstComponent, we must start with 0. 
 For intCurrentGeometrySectionIndex = 0 To intSections - 1 
 
 'Set a variable to use when accessing the current 
 'Geometry section. 
 intCurrentGeometrySection = visSectionFirstComponent + intCurrentGeometrySectionIndex 
 
 'Get the count of rows in the current Geometry section. 
 intRows = vsoShape.RowCount(intCurrentGeometrySection) 
 
 'Loop through the rows. The count is zero-based. 
 For intCurrentRow = 0 To (intRows - 1) 
 
 'Get the count of cells in the current row. 
 intCells = vsoShape.RowsCellCount(intCurrentGeometrySection, intCurrentRow) 
 
 'Loop through the cells. Again, this is zero-based. 
 For intCurrentCell = 0 To (intCells - 1) 
 
 'Get the cell's formula and 
 'add it to the list box. 
 UserForm1.ListBox1.AddItem _ 
 vsoShape.CellsSRC(intCurrentGeometrySection, intCurrentRow, _ 
 intCurrentCell).LocalName &; ": " &; _ 
 vsoShape.CellsSRC(intCurrentGeometrySection, intCurrentRow, _ 
 intCurrentCell).Formula 
 Next intCurrentCell 
 Next intCurrentRow 
 Next intCurrentGeometrySectionIndex 
 
 'Display the user form. 
 
 UserForm1.Show 
 
End Sub
```


