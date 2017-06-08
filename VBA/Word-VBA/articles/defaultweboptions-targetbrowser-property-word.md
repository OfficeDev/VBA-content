---
title: DefaultWebOptions.TargetBrowser Property (Word)
keywords: vbawd10.chm165871633
f1_keywords:
- vbawd10.chm165871633
ms.prod: word
api_name:
- Word.DefaultWebOptions.TargetBrowser
ms.assetid: e5d31e0c-d669-4b16-bf8d-0c5353732b17
ms.date: 06/08/2017
---


# DefaultWebOptions.TargetBrowser Property (Word)

Sets or returns an  **MsoTargetBrowser** constant representing the target browser for documents viewed in a Web browser. Read/write.


## Syntax

 _expression_ . **TargetBrowser**

 _expression_ Required. A variable that represents a **[DefaultWebOptions](defaultweboptions-object-word.md)** collection.


## Example

This example sets the target browser for all documents to Internet Explorer 6.


```vb
Sub GlobalTargetBrowser() 
 Application.DefaultWebOptions _ 
 .TargetBrowser = msoTargetBrowserIE6 
End Sub
```


## See also


#### Concepts


[DefaultWebOptions Object](defaultweboptions-object-word.md)

