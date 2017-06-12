---
title: Comments Object (Word)
ms.prod: word
ms.assetid: e384b37a-50e3-a214-52a8-6fda2acc4991
ms.date: 06/08/2017
---


# Comments Object (Word)

A collection of  **[Comment](comment-object-word.md)** objects that represent the comments in a selection, range, or document.


## Remarks

Use the  **Comments** property to return the **Comments** collection. The following example displays comments made by Don Funk in the active document.


```
ActiveDocument.ActiveWindow.View.SplitSpecial = wdPaneComments 
ActiveDocument.Comments.ShowBy = "Don Funk"
```

Use the  **[Add](comments-add-method-word.md)** method to add a comment at the specified range. The following example adds a comment immediately after the selection.




```
Selection.Collapse Direction:=wdCollapseEnd 
ActiveDocument.Comments.Add Range:=Selection.Range, _ 
 Text:="review this"
```

Use  **Comments** (Index), where Index is the index number, to return a single **Comment** object. The index number represents the position of the comment in the specified selection, range, or document. The following example displays the author of the first comment in the active document.




```
MsgBox ActiveDocument.Comments(1).Author
```

The following example displays the initials of the author of the first comment in the selection.




```
If Selection.Comments.Count >= 1 Then MsgBox _ 
 Selection.Comments(1).Initial
```


## Methods



|**Name**|
|:-----|
|[Add](comments-add-method-word.md)|
|[Item](comments-item-method-word.md)|

## Properties



|**Name**|
|:-----|
|[Application](comments-application-property-word.md)|
|[Count](comments-count-property-word.md)|
|[Creator](comments-creator-property-word.md)|
|[Parent](comments-parent-property-word.md)|
|[ShowBy](comments-showby-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
