---
title: XSLTransform.Alias Property (Word)
keywords: vbawd10.chm76742658
f1_keywords:
- vbawd10.chm76742658
ms.prod: word
api_name:
- Word.XSLTransform.Alias
ms.assetid: 38615e8f-cb40-6e83-f29c-520430f16ada
ms.date: 06/08/2017
---


# XSLTransform.Alias Property (Word)

Returns a  **String** that represents the display name for the specified object.


## Syntax

 _expression_ . **Alias**

 _expression_ Required. A variable that represents a **[XSLTransform](xsltransform-object-word.md)** object.


## Example

The following example shows the display name for the first schema attached to the active document.


```vb
MsgBox Application.XMLNamespaces(1).Alias
```


## See also


#### Concepts


[XSLTransform Object](xsltransform-object-word.md)

