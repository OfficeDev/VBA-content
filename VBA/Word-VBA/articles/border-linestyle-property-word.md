---
title: Border.LineStyle Property (Word)
keywords: vbawd10.chm154861571
f1_keywords:
- vbawd10.chm154861571
ms.prod: word
api_name:
- Word.Border.LineStyle
ms.assetid: 1e95d9b9-1293-753a-efbd-8fc95e9dd8b0
ms.date: 06/08/2017
---


# Border.LineStyle Property (Word)

Returns or sets the border line style for the specified object. Read/write  **WdLineStyle** .


## Syntax

 _expression_ . **LineStyle**

 _expression_ Required. A variable that represents a **[Border](border-object-word.md)** object.


## Remarks

Setting the  **LineStyle** property for a range that refers to individual characters or words applies a character border.

Setting the  **LineStyle** property for a paragraph or range of paragraphs applies a paragraph border. Use the **InsideLineStyle** property to apply a border between consecutive paragraphs.

Setting the  **LineStyle** property for a section applies a page border around the pages in the section.


## Example

If the selection is a paragraph or a collapsed selection, this example adds a single 0.75-point paragraph border above the selection. If the selection doesn't include a paragraph, a border is applied around the selected text.


```vb
With Selection.Borders(wdBorderTop) 
 .LineStyle = wdLineStyleSingle 
 .LineWidth = wdLineWidth075pt 
End With
```

This example adds a double 1.5-point border below each frame in the active document.




```vb
For Each aFrame In ActiveDocument.Frames 
 With aFrame.Borders(wdBorderBottom) 
 .LineStyle = wdLineStyleDouble 
 .LineWidth = wdLineWidth150pt 
 End With 
Next aFrame
```

The following example applies a border around the fourth word in the active document. Applying a single border (in this example, a top border) to text applies a border around the text.




```vb
ActiveDocument.Words(4).Borders(wdBorderTop) _ 
 .LineStyle = wdLineStyleSingle
```


## See also


#### Concepts


[Border Object](border-object-word.md)

