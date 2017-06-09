---
title: Application.PutFocusInMailHeader Method (Word)
keywords: vbawd10.chm158335440
f1_keywords:
- vbawd10.chm158335440
ms.prod: word
api_name:
- Word.Application.PutFocusInMailHeader
ms.assetid: ca57a93b-1487-d19c-34c9-02484ce25485
ms.date: 06/08/2017
---


# Application.PutFocusInMailHeader Method (Word)

Places the insertion point in the  **To**line of the mail header if the document in the active window is an e-mail document.


## Syntax

 _expression_ . **PutFocusInMailHeader**

 _expression_ Required. A variable that represents an **[Application](application-object-word.md)** object.


## Remarks

For best results, use the  **PutFocusInMailHeader** method with the **EnvelopeVisible** property. When the **EnvelopeVisible** property is set to **True** , the **PutFocusInMailHeader** method will place the insertion point in the mail header.


## Example

The following example displays the mail header for the active document and then place the insertion point in the  **To**line of the mail header.


```vb
ActiveDocument.ActiveWindow.EnvelopeVisible = True 
Application.PutFocusInMailHeader
```


## See also


#### Concepts


[Application Object](application-object-word.md)

