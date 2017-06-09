---
title: Application.XMLNamespaces Property (Word)
keywords: vbawd10.chm158335439
f1_keywords:
- vbawd10.chm158335439
ms.prod: word
api_name:
- Word.Application.XMLNamespaces
ms.assetid: e7eac332-f805-5ceb-682c-482565ff0786
ms.date: 06/08/2017
---


# Application.XMLNamespaces Property (Word)

Returns an  **** collection that represents the XML schemas in the Schema Library.


## Syntax

 _expression_ . **XMLNamespaces**

 _expression_ An expression that returns an **[Application](application-object-word.md)** object.


## Example

The following example returns the first schema in the Schema Library.


```vb
Dim objSchema As XMLNamespace 
 
Set objSchema = Application.XMLNamespaces.Item(1)
```


## See also


#### Concepts


[Application Object](application-object-word.md)

