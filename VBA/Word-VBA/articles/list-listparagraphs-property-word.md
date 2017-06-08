---
title: List.ListParagraphs Property (Word)
keywords: vbawd10.chm160563202
f1_keywords:
- vbawd10.chm160563202
ms.prod: word
api_name:
- Word.List.ListParagraphs
ms.assetid: 3360f8dd-155a-3b44-1b0c-395ddbac2b51
ms.date: 06/08/2017
---


# List.ListParagraphs Property (Word)

Returns a  **[ListParagraphs](listparagraphs-object-word.md)** collection that represents all the numbered paragraphs in the list, document, or range. Read-only.


## Syntax

 _expression_ . **ListParagraphs**

 _expression_ A variable that represents a **[List](list-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example double underlines the paragraphs in the second list in the active document.


```vb
For Each mypara In ActiveDocument.Lists(2).ListParagraphs 
 mypara.Range.Underline = wdUnderlineDouble 
Next mypara
```


## See also


#### Concepts


[List Object](list-object-word.md)

