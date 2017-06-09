---
title: DropCap.Clear Method (Word)
keywords: vbawd10.chm156631140
f1_keywords:
- vbawd10.chm156631140
ms.prod: word
api_name:
- Word.DropCap.Clear
ms.assetid: 8d5148ff-04ad-bb4b-7d7e-76cbc01246a9
ms.date: 06/08/2017
---


# DropCap.Clear Method (Word)

Removes the dropped capital letter formatting.


## Syntax

 _expression_ . **Clear**

 _expression_ A variable that represents a **[DropCap](dropcap-object-word.md)** object.


## Example

This example removes dropped capital letter formatting from the first letter in the active document.


```vb
Set drop = ActiveDocument.Paragraphs(1).DropCap 
If Not (drop Is Nothing) Then drop.Clear
```


## See also


#### Concepts


[DropCap Object](dropcap-object-word.md)

