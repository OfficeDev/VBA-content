---
title: XMLNamespace.Alias Property (Word)
keywords: vbawd10.chm2293764
f1_keywords:
- vbawd10.chm2293764
ms.prod: word
api_name:
- Word.XMLNamespace.Alias
ms.assetid: 3c82e7c4-4ad6-c88d-b9cb-4156534b4d89
ms.date: 06/08/2017
---


# XMLNamespace.Alias Property (Word)

Returns a  **String** that represents the display name for the specified object.


## Syntax

 _expression_ . **Alias**

 _expression_ Required. A variable that represents a **[XMLNamespace](xmlnamespace-object-word.md)** object.


## Example

The following example shows the display name for the first schema attached to the active document.


```vb
MsgBox Application.XMLNamespaces(1).Alias
```


## See also


#### Concepts


[XMLNamespace Object](xmlnamespace-object-word.md)

