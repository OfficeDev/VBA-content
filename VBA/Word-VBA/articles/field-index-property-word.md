---
title: Field.Index Property (Word)
keywords: vbawd10.chm154075144
f1_keywords:
- vbawd10.chm154075144
ms.prod: word
api_name:
- Word.Field.Index
ms.assetid: 68f3f817-1415-f428-cb38-ed79aff013e2
ms.date: 06/08/2017
---


# Field.Index Property (Word)

Returns a  **Long** that represents the position of an item in a collection. Read-only.


## Syntax

 _expression_ . **Index**

 _expression_ Required. A variable that represents a **[Field](field-object-word.md)** object.


## Example

This example returns the position of the selected field in the Fields collection.


```
num = Selection.Fields(1).Index
```


## See also


#### Concepts


[Field Object](field-object-word.md)

