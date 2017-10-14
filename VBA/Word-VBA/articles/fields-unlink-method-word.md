---
title: Fields.Unlink Method (Word)
keywords: vbawd10.chm154140774
f1_keywords:
- vbawd10.chm154140774
ms.prod: word
api_name:
- Word.Fields.Unlink
ms.assetid: 18b72e38-8a03-90fc-76f0-2f4e9d768dd9
ms.date: 06/08/2017
---


# Fields.Unlink Method (Word)

Replaces all the fields in the  **Fields** collection with their most recent results.


## Syntax

 _expression_ . **Unlink**

 _expression_ Required. A variable that represents a **[Fields](fields-object-word.md)** collection.


## Remarks

When you unlink a field, the current result is converted to text or a graphic and can no longer be updated automatically. Note that some fields—such as XE (Index Entry) fields and SEQ (Sequence) fields—cannot be unlinked.


## Example

This example updates and unlinks all the fields in the first section in the active document.


```vb
With ActiveDocument.Sections(1).Range.Fields 
 .Update 
 .Unlink 
End With
```


## See also


#### Concepts


[Fields Collection Object](fields-object-word.md)

