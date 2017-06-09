---
title: WebOptions.BrowserLevel Property (Word)
keywords: vbawd10.chm165937154
f1_keywords:
- vbawd10.chm165937154
ms.prod: word
api_name:
- Word.WebOptions.BrowserLevel
ms.assetid: f753deef-cd67-918d-0fe0-af4f3d283086
ms.date: 06/08/2017
---


# WebOptions.BrowserLevel Property (Word)

Returns or sets  **WdBrowserLevel** that represents the level of Web browser at which you want to target the specified Web page. Read/write.


## Syntax

 _expression_ . **BrowserLevel**

 _expression_ Required. A variable that represents a **[WebOptions](weboptions-object-word.md)** collection.


## Remarks

This property is ignored if the  **OptimizeForBrowser** property is set to **False** .

After you set the  **BrowserLevel** property on the **DefaultWebOptions** object, the **BrowserLevel** property of any new Web pages you create in Word will be the same as the global setting.


## See also


#### Concepts


[WebOptions Object](weboptions-object-word.md)

