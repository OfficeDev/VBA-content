---
title: View.SplitSpecial Property (Word)
keywords: vbawd10.chm161808413
f1_keywords:
- vbawd10.chm161808413
ms.prod: word
api_name:
- Word.View.SplitSpecial
ms.assetid: 5ca301aa-737f-3588-9d53-176990206620
ms.date: 06/08/2017
---


# View.SplitSpecial Property (Word)

Returns or sets the active window pane. Read/write  **WdSpecialPane** .


## Syntax

 _expression_ . **SplitSpecial**

 _expression_ A variable that represents a **[View](view-object-word.md)** object.


## Example

This example displays the primary footer in a separate pane in the active window.


```vb
ActiveDocument.ActiveWindow.View.SplitSpecial = wdPanePrimaryFooter
```

This example adds a footnote to the active document and displays all the footnotes in a separate pane in the active window.




```vb
ActiveDocument.Footnotes.Add Range:=Selection.Range, _ 
 Text:="Footnote text" 
With ActiveDocument.ActiveWindow.View 
 .Type = wdNormalView 
 .SplitSpecial = wdPaneFootnotes 
End With
```


## See also


#### Concepts


[View Object](view-object-word.md)

