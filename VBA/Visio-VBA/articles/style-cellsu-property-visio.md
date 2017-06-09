---
title: Style.CellsU Property (Visio)
keywords: vis_sdr.chm11451955
f1_keywords:
- vis_sdr.chm11451955
ms.prod: visio
api_name:
- Visio.Style.CellsU
ms.assetid: 3f4aae37-cd43-5128-a8fa-ac74c5a2b5a3
ms.date: 06/08/2017
---


# Style.CellsU Property (Visio)

Returns a  **Cell** object that represents a ShapeSheet cell. Read-only.


## Syntax

 _expression_ . **CellsU**( **_localeIndependentCellName_** )

 _expression_ A variable that represents a **Style** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _localeIndependentCellName_|Required| **String**|The name of a ShapeSheet cell.|

### Return Value

Cell


## Remarks

 **CellsU** ("somestring") raises an "Unexpected end of file" exception if "somestring" does not name an actual cell. You can use the **CellExistsU** property to determine if a cell with the universal name "somestring" exists.

The cells in a shape's User-Defined Cells and Shape Data sections belong to rows whose names have been assigned by the user or a program. You can use the  **CellsU** property to access cells in named rows.

For example, if "Row_1" is the name of a row in a shape's User-Defined Cells section, you can use this statement to access the first cell in this row (the cell in column zero, which holds the name of the row): 




```
vsoCell = vsoShape.CellsU("User.Row_1")
```

You can use this statement to access the prompt cell in Row_1: 




```
vsoCell = vsoShape.CellsU("User.Row_1.Prompt")
```

Next, assume that Row_1 is in the Shape Data section instead of the User- Defined Cells section. You can use this statement to access the first cell in the row (the cell in column zero, which holds the name of the row): 




```
vsoCell = vsoShape.CellsU("Prop.Row_1")
```

You can use this statement to access other cells in the row: 




```
vsoCell = vsoShape.CellsU("Prop.Row_1.xxx ")
```

where  _xxx_ is one of these cells: Label, Prompt, SortKey, Type, Format, Invisible, or Ask.




 **Note**  Beginning with Microsoft Visio 2000, you can use both local and universal names to refer to Visio shapes, masters, documents, pages, rows, add-ons, cells, hyperlinks, styles, fonts, master shortcuts, UI objects, and layers. When a user names a shape, for example, the user is specifying a local name. Beginning with Microsoft Office Visio 2003, the ShapeSheet spreadsheet displays only universal names in cell formulas and values. (In prior versions, universal names were not visible in the user interface.) 

As a developer, you can use universal names in a program when you don't want to change a name each time a solution is localized. Use the  **Cells** property to get a **Cell** object by using the cell's local name. Use the **CellsU** property to get a **Cell** object by using the cell's universal name.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **CellsU** property to get a particular ShapeSheet cell by its universal name. It draws a rectangle on a page and bows, or curves the lines of the rectangle by changing the shape's lines to arcs. This is accomplished by changing the ShapeSheet row types for each side of the rectangle from LineTo to ArcTo and then changing the values of the X and Y cells in each of these rows.


```vb
 
Public Sub CellsU_Example() 
 
 Dim vsoPage As Visio.Page 
 Dim vsoShape As Visio.Shape 
 Dim vsoCell As Visio.Cell 
 Dim strBowCell As String 
 Dim strBowFormula As String 
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
 
 'Add a scratch section and add a row to the scratch section. 
 vsoShape.AddSection visSectionScratch 
 vsoShape.AddRow visSectionScratch, visRowScratch, 0 
 
 'Set vsoCell to the Scratch.X1 cell and set its formula. 
 Set vsoCell = vsoShape.CellsU(strBowCell) 
 vsoCell.Formula = strBowFormula 
 
 'Bow in or curve the rectangle's lines by changing 
 'each row type from LineTo to ArcTo and entering the bow value. 
 For intCounter = 1 To 4 
 vsoShape.RowType(visSectionFirstComponent, visRowVertex + intCounter) = visTagArcTo 
 Set vsoCell = vsoShape.CellsSRC(visSectionFirstComponent, visRowVertex + intCounter, 2) 
 vsoCell.Formula = "-" &; strBowCell 
 Next intCounter 
 
End Sub
```


