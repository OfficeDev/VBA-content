---
title: WebOptions.TargetBrowser Property (Word)
keywords: vbawd10.chm165937164
f1_keywords:
- vbawd10.chm165937164
ms.prod: word
api_name:
- Word.WebOptions.TargetBrowser
ms.assetid: d503e040-9534-e3ff-a526-2ede6f595982
ms.date: 06/08/2017
---


# WebOptions.TargetBrowser Property (Word)

Sets or returns an  **MsoTargetBrowser** constant representing the target browser for documents viewed in a Web browser. Read/write.


## Syntax

 _expression_ . **TargetBrowser**

 _expression_ Required. A variable that represents a **[WebOptions](weboptions-object-word.md)** collection.


## Example

This example sets the target browser for the active document to Microsoft Internet Explorer 6 if the current target browser is an earlier version.


```vb
Sub SetWebBrowser() 
 With ActiveDocument.WebOptions 
 If .TargetBrowser < msoTargetBrowserIE6 Then 
 .TargetBrowser = msoTargetBrowserIE6 
 End If 
 End With 
End Sub
```


## See also


#### Concepts


[WebOptions Object](weboptions-object-word.md)

