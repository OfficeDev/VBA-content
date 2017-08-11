---
title: Document.ListTemplates Property (Word)
keywords: vbawd10.chm158007359
f1_keywords:
- vbawd10.chm158007359
ms.prod: word
api_name:
- Word.Document.ListTemplates
ms.assetid: dc27553a-7083-4f14-ffd6-0f440982a79c
ms.date: 06/08/2017
---


# Document.ListTemplates Property (Word)

Returns a  **ListTemplates** collection that represents all the list formats for the specified document. Read-only.


## Syntax

 _expression_ . **ListTemplates**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx). The ListTemplates property is a member of the [Document](document-object-word.md), [ListGallery](listgallery-object-word.md), and [Template](template-object-word.md) objects.


## Example

This example displays the number of list templates used in the active document.


```
Msgbox ActiveDocument.ListTemplates.Count
```


## See also


#### Concepts


[Document Object](document-object-word.md)

