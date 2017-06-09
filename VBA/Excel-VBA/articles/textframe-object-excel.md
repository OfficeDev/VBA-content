---
title: TextFrame Object (Excel)
keywords: vbaxl10.chm643072
f1_keywords:
- vbaxl10.chm643072
ms.prod: excel
api_name:
- Excel.TextFrame
ms.assetid: 4a6d2201-84b8-d83a-cc13-703da047815e
ms.date: 06/08/2017
---


# TextFrame Object (Excel)

Represents the text frame in a  **[Shape](shape-object-excel.md)** object. Contains the text in the text frame as well as the properties and methods that control the alignment and anchoring of the text frame.


## Remarks

Use the  **[TextFrame](shape-textframe-property-excel.md)** property to return a **TextFrame** object.


## Example

 The following example adds a rectangle to _myDocument_, adds text to the rectangle, and then sets the margins for the text frame.


```
Set myDocument = Worksheets(1) 
With myDocument.Shapes.AddShape(msoShapeRectangle, _ 
 0, 0, 250, 140).TextFrame 
 .Characters.Text = "Here is some test text" 
 .MarginBottom = 10 
 .MarginLeft = 10 
 .MarginRight = 10 
 .MarginTop = 10 
End With
```


## Methods



|**Name**|
|:-----|
|[Characters](textframe-characters-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](textframe-application-property-excel.md)|
|[AutoMargins](textframe-automargins-property-excel.md)|
|[AutoSize](textframe-autosize-property-excel.md)|
|[Creator](textframe-creator-property-excel.md)|
|[HorizontalAlignment](textframe-horizontalalignment-property-excel.md)|
|[HorizontalOverflow](textframe-horizontaloverflow-property-excel.md)|
|[MarginBottom](textframe-marginbottom-property-excel.md)|
|[MarginLeft](textframe-marginleft-property-excel.md)|
|[MarginRight](textframe-marginright-property-excel.md)|
|[MarginTop](textframe-margintop-property-excel.md)|
|[Orientation](textframe-orientation-property-excel.md)|
|[Parent](textframe-parent-property-excel.md)|
|[ReadingOrder](textframe-readingorder-property-excel.md)|
|[VerticalAlignment](textframe-verticalalignment-property-excel.md)|
|[VerticalOverflow](textframe-verticaloverflow-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
