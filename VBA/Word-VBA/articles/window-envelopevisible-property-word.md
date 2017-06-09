---
title: Window.EnvelopeVisible Property (Word)
keywords: vbawd10.chm157417505
f1_keywords:
- vbawd10.chm157417505
ms.prod: word
api_name:
- Word.Window.EnvelopeVisible
ms.assetid: d04d6714-ba32-39cc-4853-e9ac6696e718
ms.date: 06/08/2017
---


# Window.EnvelopeVisible Property (Word)

 **True** if the e-mail message header is visible in the document window. The default value is **False** . Read/write **Boolean** .


## Syntax

 _expression_ . **EnvelopeVisible**

 _expression_ A variable that represents a **[Window](window-object-word.md)** object.


## Remarks

This property has no effect if the document isn't an e-mail message.


## Example

This example displays the e-mail message header.


```vb
ActiveWindow.EnvelopeVisible = True
```


## See also


#### Concepts


[Window Object](window-object-word.md)

