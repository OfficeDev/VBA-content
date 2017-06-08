---
title: Style.Table Property (Word)
keywords: vbawd10.chm153878549
f1_keywords:
- vbawd10.chm153878549
ms.prod: word
api_name:
- Word.Style.Table
ms.assetid: bc078a71-232f-b2c0-e9be-df9bda492e5e
ms.date: 06/08/2017
---


# Style.Table Property (Word)

Returns a  **[TableStyle](tablestyle-object-word.md)** object representing properties that can be applied to a table using a table style.


## Syntax

 _expression_ . **Table**

 _expression_ An expression that returns a **[Style](style-object-word.md)** object.


## Example

This example creates a new table style that specifies a surrounding border and special borders and shading for only the first and last rows and the last column.


```vb
Sub NewTableStyle() 
 Dim styTable As Style 
 
 Set styTable = ActiveDocument.Styles.Add( _ 
 Name:="TableStyle 1", Type:=wdStyleTypeTable) 
 
 With styTable.Table 
 
 'Apply borders around table, a double border to the heading row, 
 'a double border to the last column, and shading to last row 
 .Borders(wdBorderTop).LineStyle = wdLineStyleSingle 
 .Borders(wdBorderBottom).LineStyle = wdLineStyleSingle 
 .Borders(wdBorderLeft).LineStyle = wdLineStyleSingle 
 .Borders(wdBorderRight).LineStyle = wdLineStyleSingle 
 
 .Condition(wdFirstRow).Borders(wdBorderBottom) _ 
 .LineStyle = wdLineStyleDouble 
 
 .Condition(wdLastColumn).Borders(wdBorderLeft) _ 
 .LineStyle = wdLineStyleDouble 
 
 .Condition(wdLastRow).Shading _ 
 .BackgroundPatternColor = wdColorGray125 
 
 End With 
End Sub
```


## See also


#### Concepts


[Style Object](style-object-word.md)

