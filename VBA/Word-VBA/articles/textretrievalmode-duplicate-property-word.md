---
title: TextRetrievalMode.Duplicate Property (Word)
keywords: vbawd10.chm154730497
f1_keywords:
- vbawd10.chm154730497
ms.prod: word
api_name:
- Word.TextRetrievalMode.Duplicate
ms.assetid: 3ccc1c6a-c709-ea9a-052d-a5c3d566038f
ms.date: 06/08/2017
---


# TextRetrievalMode.Duplicate Property (Word)

Returns a read-only  **TextRetrievalMode** object that represents options related to retrieving text from a **Range** object.


## Syntax

 _expression_ . **Duplicate**

 _expression_ Required. A variable that represents a **[TextRetrievalMode](textretrievalmode-object-word.md)** object.


## Remarks

You can use the  **Duplicate** property to pick up the settings of all the properties of a duplicated object. You can assign the object returned by the **Duplicate** property to another object of the same type to apply those settings all at once. Before assigning the duplicate object to another object, you can change any of the properties of the duplicate object without affecting the original text.


## See also


#### Concepts


[TextRetrievalMode Object](textretrievalmode-object-word.md)

