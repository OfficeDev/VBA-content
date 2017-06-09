---
title: Border.LineWidth Property (Word)
keywords: vbawd10.chm154861572
f1_keywords:
- vbawd10.chm154861572
ms.prod: word
api_name:
- Word.Border.LineWidth
ms.assetid: 31e87acf-fd7f-fa5c-d869-5f46bb7ed169
ms.date: 06/08/2017
---


# Border.LineWidth Property (Word)

Returns or sets the line width of an object's border. Read/write.


## Syntax

 _expression_ . **LineWidth**

 _expression_ Required. A variable that represents a **[Border](border-object-word.md)** object.


## Remarks

Returns a  **WdLineWidth** constant or **wdUndefined** if the object either has no borders or has borders with more than one line width. If the specified line width isn't available for the border's line style, this property generates an error. To determine the line widths available for a particular line style, see the **Borders and Shading** dialog box ( **Format** menu).


## Example

This example adds a border below the first row in the first table of the active document.


```vb
If ActiveDocument.Tables.Count >= 1 Then 
 With ActiveDocument.Tables(1).Rows(1).Borders(wdBorderBottom) 
 .LineStyle = wdLineStyleSingle 
 .LineWidth = wdLineWidth050pt 
 End With 
End If
```

This example adds a wavy, red line to the left of the selection.




```vb
With Selection.Borders(wdBorderLeft) 
 .LineStyle = wdLineStyleSingleWavy 
 .LineWidth = wdLineWidth075pt 
 .ColorIndex = wdRed 
End With
```


## See also


#### Concepts


[Border Object](border-object-word.md)

