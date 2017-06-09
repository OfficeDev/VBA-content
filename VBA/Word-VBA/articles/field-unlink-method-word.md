---
title: Field.Unlink Method (Word)
keywords: vbawd10.chm154075238
f1_keywords:
- vbawd10.chm154075238
ms.prod: word
api_name:
- Word.Field.Unlink
ms.assetid: b547d99e-fbf7-f31a-ca98-c9fab1e006e7
ms.date: 06/08/2017
---


# Field.Unlink Method (Word)

Replaces the specified field with its most recent result.


## Syntax

 _expression_ . **Unlink**

 _expression_ Required. A variable that represents a **[Field](field-object-word.md)** object.


## Remarks

When you unlink a field, the current result is converted to text or a graphic and can no longer be updated automatically. Note that some fields—such as XE (Index Entry) fields and SEQ (Sequence) fields—cannot be unlinked.


## Example

This example unlinks the first field in "Sales.doc."


```
Documents("Sales.doc").Fields(1).Unlink
```


## See also


#### Concepts


[Field Object](field-object-word.md)

