---
title: Endnote.Range Property (Word)
keywords: vbawd10.chm155058180
f1_keywords:
- vbawd10.chm155058180
ms.prod: word
api_name:
- Word.Endnote.Range
ms.assetid: fde6bb87-f2ce-7bf4-ecc3-a78b8db0e1c4
ms.date: 06/08/2017
---


# Endnote.Range Property (Word)

Returns a  **Range** object that represents the portion of a document that is contained in the specified object.


## Syntax

 _expression_ . **Range**

 _expression_ Required. A variable that represents an **[Endnote](endnote-object-word.md)** object.


## Remarks

For information about returning a range from a document or returning a shape range from a collection of shapes, see the  **Range** method.


## Example

This example changes the text of the first endnote in the active document.


```vb
With ActiveDocument.Endnotes(1).Range 
 .Delete 
 .Text = "new endnote text" 
End With
```


## See also


#### Concepts


[Endnote Object](endnote-object-word.md)

