---
title: View.SeekView Property (Word)
keywords: vbawd10.chm161808412
f1_keywords:
- vbawd10.chm161808412
ms.prod: word
api_name:
- Word.View.SeekView
ms.assetid: 94b026a0-92f9-32c4-0394-d2b02fbcb942
ms.date: 06/08/2017
---


# View.SeekView Property (Word)

Returns or sets the document element displayed in print layout view. Read/write  **WdSeekView** .


## Syntax

 _expression_ . **SeekView**

 _expression_ Required. A variable that represents a **[View](view-object-word.md)** object.


## Remarks

This property generates an error if the view is not print layout view.


## Example

If the active document has footnotes, this example displays footnotes in print layout view.


```vb
If ActiveDocument.Footnotes.Count >= 1 Then 
 With ActiveDocument.ActiveWindow.View 
 .Type = wdPrintView 
 .SeekView = wdSeekFootnotes 
 End With 
End If
```

This example shows the first page footer for the current section.




```vb
ActiveDocument.PageSetup.DifferentFirstPageHeaderFooter = True 
With ActiveDocument.ActiveWindow.View 
 .Type = wdPrintView 
 .SeekView = wdSeekFirstPageFooter 
End With
```

If the selection is in a footnote or endnote area in print layout view, this example switches to the main document.




```vb
Set myView = ActiveDocument.ActiveWindow.View 
If myView.SeekView = wdSeekFootnotes Or _ 
 myView.SeekView = wdSeekEndnotes Then 
 myView.SeekView = wdSeekMainDocument 
End If
```


## See also


#### Concepts


[View Object](view-object-word.md)

