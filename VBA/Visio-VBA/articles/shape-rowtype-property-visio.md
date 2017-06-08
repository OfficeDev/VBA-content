---
title: Shape.RowType Property (Visio)
keywords: vis_sdr.chm11214270
f1_keywords:
- vis_sdr.chm11214270
ms.prod: visio
api_name:
- Visio.RowType
ms.assetid: 416b77f1-6cec-de5b-c2b8-c6e5b239c54c
ms.date: 06/08/2017
---


# Shape.RowType Property (Visio)

Gets or sets the type of a row in a Geometry, Connection Points, Controls, or Tabs ShapeSheet section. Read/write.


## Syntax

 _expression_ . **RowType**( **_Section_** , **_Row_** )

 _expression_ A variable that represents a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Section_|Required| **Integer**|The index of the section that contains the row.|
| _Row_|Required| **Integer**|The index of the row.|

### Return Value

Integer


## Remarks

After you change a row's type, the new row type may or may not have the same cells. Your program must provide the appropriate formulas for the new or changed cells.

You can specify the type of row you want by setting  **RowType** equal to any of the following constants declared by the Visio type library in member **[VisRowTags](visrowtags-enumeration-visio.md)** .



|**Constant**|**Value**|
|:-----|:-----|
| **visTagComponent**|137|
| **visTagMoveTo**|138|
| **visTagLineTo**|139|
| **visTagArcTo**|140|
| **visTagInfiniteLine**|141|
| **visTagEllipse**|143|
| **visTagEllipticalArcTo**|144|
| **visTagSplineBeg**|165|
| **visTagSplineSpan**|166|
| **visTagPolylineTo**|193|
| **visTagNURBSTo**|195|
| **visTagTab0**|136|
| **visTagTab2**|150|
| **visTagTab10**|151|
| **visTagTab60**|181|
| **visTagCnnctPt**|153|
| **visTagCnnctNamed**|185|
| **visTagCtlPt**|162|
| **visTagCtlPtTip**|170|
If an inappropriate row tag is passed or the row does not exist, no changes occur and an error is returned.

Use the  **RowName** property to transition from unnamed to named Connection Points rows.

See  **[VisRowIndices](visrowindices-enumeration-visio.md)** for a list of valid row constants and **[VisRowTags](visrowtags-enumeration-visio.md)** for a list of valid row tag constants.

See  **[VisSectionIndices](vissectionindices-enumeration-visio.md)** for a list of valid section constants.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **RowType** property to change the type of a ShapeSheet row. It draws a rectangle on a page and bows, or curves, the lines of the rectangle by changing the shape's lines to arcs. This is accomplished by changing the ShapeSheet row types for each side of the rectangle from LineTo to ArcTo and then changing the values of the X and Y cells in each of these rows.


```vb
 
Public Sub RowType_Example() 
 
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
 Set vsoCell = vsoShape.Cells(strBowCell) 
 vsoCell.formula = strBowFormula 
 
 'Bow in or curve the rectangle's lines by changing 
 'each row type from LineTo to ArcTo and entering the bow value. 
 For intCounter = 1 To 4 
 
 vsoShape.RowType(visSectionFirstComponent, visRowVertex + intCounter) = visTagArcTo 
 Set vsoCell = vsoShape.CellsSRC(visSectionFirstComponent, visRowVertex + intCounter, 2) 
 vsoCell.formula = "-" &; strBowCell 
 
 Next intCounter 
 
End Sub
```


