---
title: RevisionsFilter.View Property (Word)
keywords: vbawd10.chm40566785
f1_keywords:
- vbawd10.chm40566785
ms.prod: word
ms.assetid: b433594a-927c-57fd-a7fd-82f8c752864e
ms.date: 06/08/2017
---


# RevisionsFilter.View Property (Word)

Sets or returns a [WdRevisionsView](wdrevisionsview-enumeration-word.md) constant that represents the global option that specifies whether Word displays the original version of a document or the final version, which might have revisions and formatting changes applied. Read/write.


## Syntax

 _expression_ . **View**

 _expression_ A variable that represents a **RevisionsFilter** object.


## Remarks

The  **RevisionsFilter.View** property replaces the **View.RevisionsView** property that was in previous version of Word.


## Example

This example toggles between displaying the original and the final version of the document. This example assumes that the document in the active window contains revisions made by one or more reviewers.


```vb
Sub RevisionsFilter_View_Example()

    With ActiveWindow.View

       If .RevisionsFilter.View = wdRevisionsViewFinal Then
                .RevisionsFilter.View = wdRevisionsViewOriginal
            Else
                .RevisionsFilter.View = wdRevisionsViewFinal
       End If
    End With
End Sub
```


## Property value

 **WDREVISIONSVIEW**


## See also


#### Other resources


[RevisionsFilter Object](revisionsfilter-object-word.md)


