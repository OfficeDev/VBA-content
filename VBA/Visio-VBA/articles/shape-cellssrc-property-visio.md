---
title: Shape.CellsSRC Property (Visio)
keywords: vis_sdr.chm11213205
f1_keywords:
- vis_sdr.chm11213205
ms.prod: visio
api_name:
- Visio.Shape.CellsSRC
ms.assetid: 8fb6fd7b-e0ca-c694-3b9d-5390d4192565
ms.date: 06/08/2017
---


# Shape.CellsSRC Property (Visio)

Returns a  **Cell** object that represents a ShapeSheet cell identified by section, row, and column indices. Read-only.


## Syntax

 _expression_ . **CellsSRC**( **_Section_** , **_Row_** , **_Column_** )

 _expression_ A variable that represents a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Section_|Required| **Integer**|The cell's section index.|
| _Row_|Required| **Integer**|The cell's row index.|
| _Column_|Required| **Integer**|The cell's column index.|

### Return Value

Cell


## Remarks

To access any shape formula by its section, row, and column indices, use the  **CellsSRC** property. Constants for section, row, and column indices are declared by the Visio type library as members of **[VisSectionIndices](vissectionindices-enumeration-visio.md)** , **[VisRowIndices](visrowindices-enumeration-visio.md)** , and **[VisCellIndices](viscellindices-enumeration-visio.md)** , respectively.

The  **CellsSRC** property might raise an exception if index values for section, row, and column do not identify an actual cell, depending on the section. However, even if no exception is raised, subsequent methods invoked on the returned object fail. You can determine if a cell with particular index values exists by using the **CellsSRCExists** property.

The  **CellsSRC** property is typically used to iterate through the cells in a section or row. To retrieve a single cell, use the **Cells** property and specify a cell name. For example:




```vb
Set vsoCell = Cells("PinX")
```

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this property maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVShape.get_CellsSRC**
    

## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **CellsSRC** property to set a particular ShapeSheet cell by its section, row, and column indices. It draws a rectangle on a page and bows, or curves the lines of the rectangle by changing the shape's lines to arcs. The macro then draws an inner rectangle within the bowed lines of the first rectangle.


```vb
 
Public Sub CellsSRC_Example() 
  
    Dim vsoPage As Visio.Page  
    Dim vsoShape As Visio.Shape  
    Dim vsoCell As Visio.Cell  
    Dim strBowCell As String 
    Dim strBowFormula As String 
    Dim intIndex As Integer 
    Dim intCounter As Integer 
 
    'Set the value of the strBowCell string.  
    strBowCell = "Scratch.X1"  
 
    'Set the value of the strBowFormula string.  
    strBowFormula = "=Min(Width, Height) / 5"  
 
    Set vsoPage = ActivePage  
 
    'If there isn't an active page, set vsoPage 
    'to the first page of the active document. 
    If vsoPage Is Nothing Then 
        Set vsoPage = ActiveDocument.Pages(1)  
    End If   
 
    'Draw a rectangle on the active page. 
    Set vsoShape = vsoPage.DrawRectangle(1, 5, 5, 1)  
 
    'Add a scratch section to the shape's ShapeSheet  
    vsoShape.AddSection visSectionScratch  
 
    'Add a row to the scratch section.  
    vsoShape.AddRow visSectionScratch, visRowScratch, 0  
 
    'Set vsoCell to the Scratch.X1 cell and set its formula. 
    Set vsoCell = vsoShape.Cells(strBowCell)  
    vsoCell.Formula = strBowFormula  
 
    'Bow in or curve the rectangle's lines by changing 
    'each row type from LineTo to ArcTo and entering the bow value. 
    For intCounter = 1 To 4  
        vsoShape.RowType(visSectionFirstComponent, visRowVertex + intCounter) = visTagArcTo  
        Set vsoCell = vsoShape.CellsSRC(visSectionFirstComponent, visRowVertex + intCounter, 2)  
        vsoCell.Formula = "-" &; strBowCell  
    Next intCounter  
 
    'Create an inner rectangle. 
    'Set the section index for the inner rectangle's Geometry section.  
    intIndex = visSectionFirstComponent + 1  
 
    'Add an inner rectangle Geometry section.  
    vsoShape.AddSection intIndex  
 
    'Add the first 2 rows to the section.  
    vsoShape.AddRow intIndex, visRowComponent, visTagComponent  
    vsoShape.AddRow intIndex, visRowVertex, visTagMoveTo 
  
    'Add 4 LineTo rows to the section 
    For intCounter = 1 To 4  
        vsoShape.AddRow intIndex, visRowLast, visTagLineTo  
    Next intCounter  
 
    'Set the inner rectangle start point cell formulas. 
    Set vsoCell = vsoShape.CellsSRC(intIndex, 1, 0)  
    vsoCell.Formula = "Width * 0 + " &; strBowCell  
    Set vsoCell = vsoShape.CellsSRC(intIndex, 1, 1)  
    vsoCell.Formula = "Height * 0 + " &; strBowCell  
 
    'Draw the inner rectangle bottom line. 
    Set vsoCell = vsoShape.CellsSRC(intIndex, 2, 0)  
    vsoCell.Formula = "Width * 1 - " &; strBowCell  
    Set vsoCell = vsoShape.CellsSRC(intIndex, 2, 1)  
    vsoCell.Formula = "Height * 0 + " &; strBowCell 
  
    'Draw the inner rectangle right side line. 
    Set vsoCell = vsoShape.CellsSRC(intIndex, 3, 0)  
    vsoCell.Formula = "Width * 1 - " &; strBowCell  
    Set vsoCell = vsoShape.CellsSRC(intIndex, 3, 1)  
    vsoCell.Formula = "Height * 1 - " &; strBowCell  
 
    'Draw the inner rectangle top line. 
    Set vsoCell = vsoShape.CellsSRC(intIndex, 4, 0)  
    vsoCell.Formula = "Width * 0 + " &; strBowCell  
    Set vsoCell = vsoShape.CellsSRC(intIndex, 4, 1)  
    vsoCell.Formula = "Height * 1 - " &; strBowCell  
 
    'Draw the inner rectangle left side line. 
    Set vsoCell = vsoShape.CellsSRC(intIndex, 5, 0)  
    vsoCell.Formula = "Geometry2.X1"  
    Set vsoCell = vsoShape.CellsSRC(intIndex, 5, 1)  
    vsoCell.Formula = "Geometry2.Y1"  
 
End Sub
```


