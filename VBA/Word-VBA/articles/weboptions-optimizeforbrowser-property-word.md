---
title: WebOptions.OptimizeForBrowser Property (Word)
keywords: vbawd10.chm165937153
f1_keywords:
- vbawd10.chm165937153
ms.prod: word
api_name:
- Word.WebOptions.OptimizeForBrowser
ms.assetid: c7b9f987-d13e-a95d-e40d-3b1c9b7f9fa0
ms.date: 06/08/2017
---


# WebOptions.OptimizeForBrowser Property (Word)

 **True** if Word optimizes the specified Web page for the Web browser specified by the **[BrowserLevel](weboptions-browserlevel-property-word.md)** property. Read/write **Boolean** .


## Syntax

 _expression_ . **OptimizeForBrowser**

 _expression_ Required. A variable that represents a **[WebOptions](weboptions-object-word.md)** collection.


## Example

This example creates a new Web page and optimizes it for Microsoft Internet Explorer 5.


```vb
Documents.Add DocumentType:=wdNewWebPage 
With ActiveDocument.WebOptions 
 .BrowserLevel = wdBrowserLevelMicrosoftInternetExplorer5 
 .OptimizeForBrowser = True 
End With
```


## See also


#### Concepts


[WebOptions Object](weboptions-object-word.md)

