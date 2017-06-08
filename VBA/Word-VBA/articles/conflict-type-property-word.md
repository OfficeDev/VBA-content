---
title: Conflict.Type Property (Word)
keywords: vbawd10.chm78708740
f1_keywords:
- vbawd10.chm78708740
ms.prod: word
api_name:
- Word.Conflict.Type
ms.assetid: d2e5ad43-4b4b-8ce2-3aeb-453012759d9a
ms.date: 06/08/2017
---


# Conflict.Type Property (Word)

Returns the [WdRevisionType](wdrevisiontype-enumeration-word.md)for the [Conflict](conflict-object-word.md) object. Read-only.


## Syntax

 _expression_ . **Type**

 _expression_ An expression that returns a **[Conflict](conflict-object-word.md)** object.


## Example

The following code example gets the [type](conflict-type-property-word.md) of each conflict in the active document.


```vb
Dim con as Conflict 
 
For Each con in ActiveDocument.CoAuthoring.Conflicts 
 MsgBox con.Type 
Next con
```


## See also


#### Concepts


[Conflict Object](conflict-object-word.md)

