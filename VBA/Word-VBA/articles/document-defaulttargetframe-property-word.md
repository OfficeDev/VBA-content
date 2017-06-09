---
title: Document.DefaultTargetFrame Property (Word)
keywords: vbawd10.chm158007636
f1_keywords:
- vbawd10.chm158007636
ms.prod: word
api_name:
- Word.Document.DefaultTargetFrame
ms.assetid: 4439bf14-34da-62b6-a290-f374eeef908a
ms.date: 06/08/2017
---


# Document.DefaultTargetFrame Property (Word)

Returns or sets a  **String** indicating the browser frame in which to display a Web page reached through a hyperlink. Read/write.


## Syntax

 _expression_ . **DefaultTargetFrame**

 _expression_ An expression that returns a **[Document](document-object-word.md)** object.


## Remarks

While the  **DefaultTargetFrame** property can use any user-defined string, it has the following predefined strings: "_top", "_blank", "_parent", and "_self".


## Example

This example sets Microsoft Word to open a new blank browser window when a user clicks a hyperlink in the active document.


```vb
Sub DefaultFrame() 
 ActiveDocument.DefaultTargetFrame = "_blank" 
End Sub
```


## See also


#### Concepts


[Document Object](document-object-word.md)

