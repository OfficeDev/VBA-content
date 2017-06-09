---
title: Cell.Dependents Property (Visio)
keywords: vis_sdr.chm10151825
f1_keywords:
- vis_sdr.chm10151825
ms.prod: visio
api_name:
- Visio.Cell.Dependents
ms.assetid: 99a1502b-c847-6836-2470-178b595345f9
ms.date: 06/08/2017
---


# Cell.Dependents Property (Visio)

Returns an array of ShapeSheet cells that are dependent on a particular cell of a Microsoft Visio shape. Read-only.


## Syntax

 _expression_ . **Dependents**

 _expression_ An expression that returns a **Cell** object.


### Return Value

Cell()


## Remarks

The  **Dependents** property returns an array of the cells that recalculate their values when the formula or value of the parent **Cell** object changes.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Dependents** property to display a list of cells that are dependent on the Width cell of a rectangle shape.


```vb
Public Sub Dependents_Example() 
 
 Dim acellDependentCells() As Visio.Cell 
 Dim vsoCell As Visio.Cell 
 Dim vsoShape As Visio.Shape 
 Dim intCounter As Integer 
 
 'Draw a rectangle on the active page. 
 Set vsoShape = ActivePage.DrawRectangle(1, 5, 5, 1) 
 
 'Get the array of cells dependent on the Width cell of the shape 
 acellDependentCells = vsoShape.Cells("Width").Dependents 
 
 'List the cell names and their associated formulas 
 For intCounter = LBound(acellDependentCells) To UBound(acellDependentCells) 
 
 Set vsoCell = acellDependentCells(intCounter) 
 Debug.Print vsoCell.Name &; " has this formula: " &; vsoCell.Formula 
 
 Next 
 
End Sub
```


