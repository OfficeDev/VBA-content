---
title: Comment Object (Word)
keywords: vbawd10.chm2365
f1_keywords:
- vbawd10.chm2365
ms.prod: word
api_name:
- Word.Comment
ms.assetid: 0a2841f3-ca3c-8186-afab-f634ebd97d4c
ms.date: 06/08/2017
---


# Comment Object (Word)

Represents a single comment. The  **Comment** object is a member of the **[Comments](comments-object-word.md)** collection. The **Comments** collection includes comments in a selection, range or document.


## Remarks

Use  **Comments** (Index), where Index is the index number, to return a single **Comment** object. The index number represents the position of the comment in the specified selection, range, or document. The following example displays the author of the first comment in the active document.


```
MsgBox ActiveDocument.Comments(1).Author
```

Use the  **[Add](comments-add-method-word.md)** method to add a comment at the specified range. The following example adds a comment immediately after the selection.




```
Selection.Collapse Direction:=wdCollapseEnd 
ActiveDocument.Comments.Add Range:=Selection.Range, _ 
 Text:="review this"
```

Use the  **[Reference](comment-reference-property-word.md)** property to return the reference mark associated with the specified comment. Use the **[Range](comment-range-property-word.md)** property to return the text associated with the specified comment. The following example displays the text associated with the first comment in the active document.




```
MsgBox ActiveDocument.Comments(1).Range.Text
```


## Methods



|**Name**|
|:-----|
|[DeleteRecursively](comment-deleterecursively-method-word.md)|
|[Edit](comment-edit-method-word.md)|

## Properties



|**Name**|
|:-----|
|[Ancestor](comment-ancestor-property-word.md)|
|[Application](comment-application-property-word.md)|
|[Contact](comment-contact-property-word.md)|
|[Creator](comment-creator-property-word.md)|
|[Date](comment-date-property-word.md)|
|[Done](comment-done-property-word.md)|
|[Index](comment-index-property-word.md)|
|[IsInk](comment-isink-property-word.md)|
|[Parent](comment-parent-property-word.md)|
|[Range](comment-range-property-word.md)|
|[Reference](comment-reference-property-word.md)|
|[Replies](comment-replies-property-word.md)|
|[Scope](comment-scope-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
