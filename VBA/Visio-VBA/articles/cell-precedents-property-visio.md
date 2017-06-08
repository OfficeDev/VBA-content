---
title: Cell.Precedents Property (Visio)
keywords: vis_sdr.chm10151690
f1_keywords:
- vis_sdr.chm10151690
ms.prod: visio
api_name:
- Visio.Cell.Precedents
ms.assetid: 4461b45a-6fd6-4376-f8b2-4d8a9597111a
ms.date: 06/08/2017
---


# Cell.Precedents Property (Visio)

Returns an array of ShapeSheet cells upon which the formula of another cell depends. Read-only.


## Syntax

 _expression_ . **Precedents**

 _expression_ A variable that represents a **Cell** object.


### Return Value

Cell()


## Remarks

The  **Precedents** property returns an array of the cells that cause the parent **Cell** object to recalculate its value when their formula or value changes.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Precedents** property to display a list of cells upon which the "Scratch.X1" cell of a shape depend. The macro draws a rectangle on the active page, adds a Scratch section to the ShapeSheet of the rectangle, and then enters a formula in a cell of that section that is used to bow the sides of the rectangle inward, by changing each of the rectangles sides to an arc. Because the formula used to bow the sides of the rectangle depends on the width and height of the rectangle, the cell that contains the formula, Scratch.X1, is dependent upon the Width and Height cells of the rectangle shape, making these cells precedents.


```vb
Public Sub Precedents_Example() 
 
 Dim acellPrecedentCells() As Visio.Cell 
 Dim vsoCell As Visio.Cell 
 Dim vsoShape As Visio.Shape 
 Dim strBowCell As String 
 Dim strBowFormula As String 
 Dim intCounter As Integer 
 
 'Set the value of the strBowCell string 
 strBowCell = "Scratch.X1" 
 
 'Set the value of the strBowFormula string 
 strBowFormula = "=Min(Width, Height) / 5" 
 
 'Draw a rectangle on the active page 
 Set vsoShape = ActivePage.DrawRectangle(1, 5, 5, 1) 
 
 'Add a scratch section and then 
 vsoShape.AddSection visSectionScratch 
 
 'Add a row to the scratch section 
 vsoShape.AddRow visSectionScratch, visRowScratch, 0 
 
 'Place the value of strBowFormula into Scratch.X1 
 'Set the Cell object to the Scratch.X1 and set formula 
 Set vsoCell = vsoShape.Cells(strBowCell) 
 
 'Set up the offset for the arc 
 vsoCell.Formula = strBowFormula 
 
 'Bow in or curve the original rectangle's lines by changing 
 'each row to an arc and entering the bow value 
 For intCounter = 1 To 4 
 
 vsoShape.RowType(visSectionFirstComponent, visRowVertex + intCounter) = visTagArcTo 
 Set vsoCell = vsoShape.CellsSRC(visSectionFirstComponent, visRowVertex + intCounter, 2) 
 vsoCell.Formula = "-" &; strBowCell 
 
 Next intCounter 
 
 'Get the array of precedent cells 
 acellPrecedentCells = vsoShape.Cells("Scratch.X1").Precedents 
 
 'List the cell names and their associated formula 
 For intCounter = LBound(acellPrecedentCells) To UBound(acellPrecedentCells) 
 Set vsoCell = acellPrecedentCells(intCounter) 
 Debug.Print vsoCell.Name &; " has this formula: " &; vsoCell.Formula 
 Next 
 
End Sub
```


