---
title: View.PreviousHeaderFooter Method (Word)
keywords: vbawd10.chm161808489
f1_keywords:
- vbawd10.chm161808489
ms.prod: word
api_name:
- Word.View.PreviousHeaderFooter
ms.assetid: fee57f18-348d-a671-2fb2-1f9797c39727
ms.date: 06/08/2017
---


# View.PreviousHeaderFooter Method (Word)

Moves to the previous header or footer, depending on whether a header or footer is displayed in the view.


## Syntax

 _expression_ . **PreviousHeaderFooter**

 _expression_ Required. A variable that represents a **[View](view-object-word.md)** object.


## Remarks

If the view displays a header, this method moves to the previous header within the current section (for example, from an even header to an odd header) or to the last header in the previous section. If the view displays a footer, this method moves to the previous footer.


 **Note**  If the view displays the first header or footer in the first section of the document, or if it is not displaying a header or footer at all, an error occurs.


## Example

This example inserts an even section break, switches the active window to print layout view, displays the current header, and then switches to the previous header.


```vb
Selection.Collapse Direction:=wdCollapseStart 
Selection.InsertBreak Type:=wdSectionBreakEvenPage 
With ActiveDocument.ActiveWindow.View 
 .Type = wdPrintView 
 .SeekView = wdSeekCurrentPageHeader 
 .PreviousHeaderFooter 
End With
```


## See also


#### Concepts


[View Object](view-object-word.md)

