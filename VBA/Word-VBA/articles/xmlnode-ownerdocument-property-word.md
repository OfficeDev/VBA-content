---
title: XMLNode.OwnerDocument Property (Word)
keywords: vbawd10.chm37748747
f1_keywords:
- vbawd10.chm37748747
ms.prod: word
api_name:
- Word.XMLNode.OwnerDocument
ms.assetid: 015559a7-6824-f8dd-edfd-d8d996ac18fc
ms.date: 06/08/2017
---


# XMLNode.OwnerDocument Property (Word)

Returns a  **Document** object that represents the parent document of the specified XML element.


## Syntax

 _expression_ . **OwnerDocument**

 _expression_ An expression that returns an **[XMLNode](xmlnode-object-word.md)** object.


## Example

The following example accesses the parent document of the first XML element in the active selection.


```vb
Dim objDoc As Document 
 
Set objDoc = Selection.XMLNodes(1).OwnerDocument
```


## See also


#### Concepts


[XMLNode Object](xmlnode-object-word.md)

