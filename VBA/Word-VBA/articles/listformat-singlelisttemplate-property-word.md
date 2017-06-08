---
title: ListFormat.SingleListTemplate Property (Word)
keywords: vbawd10.chm163577929
f1_keywords:
- vbawd10.chm163577929
ms.prod: word
api_name:
- Word.ListFormat.SingleListTemplate
ms.assetid: 9f02aa2f-c855-b117-c031-d03bac3d5f53
ms.date: 06/08/2017
---


# ListFormat.SingleListTemplate Property (Word)

 **True** if the entire **ListFormat** object uses the same list template. Read-only **Boolean** .


## Syntax

 _expression_ . **SingleListTemplate**

 _expression_ Required. A variable that represents a **[ListFormat](listformat-object-word.md)** object.


## Example

This example checks to see whether the selection is formatted with a single list template. If so, the example applies the second numbered list template to the selection.


```vb
Set myList = Selection.Range.ListFormat 
temp = myList.SingleListTemplate 
If temp = True Then 
 myList.ApplyListTemplate _ 
 ListTemplate:=ListGalleries(wdNumberGallery) _ 
 .ListTemplates(2) 
End If
```


## See also


#### Concepts


[ListFormat Object](listformat-object-word.md)

