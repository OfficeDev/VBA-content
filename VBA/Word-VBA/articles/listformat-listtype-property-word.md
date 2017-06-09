---
title: ListFormat.ListType Property (Word)
keywords: vbawd10.chm163577930
f1_keywords:
- vbawd10.chm163577930
ms.prod: word
api_name:
- Word.ListFormat.ListType
ms.assetid: 6a6cf33b-d1a7-25f8-2fe0-ab98760c424e
ms.date: 06/08/2017
---


# ListFormat.ListType Property (Word)

Returns the type of lists that are contained in the range for the specified  **ListFormat** object. Read-only **WdListType** .


## Syntax

 _expression_ . **ListType**

 _expression_ Required. A variable that represents a **[ListFormat](listformat-object-word.md)** object.


## Remarks

The constant  **wdListListNumOnly** refers to LISTNUM fields, which are fields that can be added within the text of a paragraph.


## Example

This example checks to see if the first list in the active document is a simple numbered list. If it is, the fourth list template on the  **Numbered** tab of the **Bullets and Numbering** dialog box ( **Format** menu) is applied.


```vb
Set myList = ActiveDocument.Lists(1) 
If myList.Range.ListFormat.ListType = wdListSimpleNumbering Then 
 myList.ApplyListTemplate _ 
 ListTemplate:=ListGalleries(wdNumberGallery) _ 
 .ListTemplates(4) 
End If
```


## See also


#### Concepts


[ListFormat Object](listformat-object-word.md)

