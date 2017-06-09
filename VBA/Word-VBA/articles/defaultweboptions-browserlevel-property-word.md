---
title: DefaultWebOptions.BrowserLevel Property (Word)
keywords: vbawd10.chm165871618
f1_keywords:
- vbawd10.chm165871618
ms.prod: word
api_name:
- Word.DefaultWebOptions.BrowserLevel
ms.assetid: 15817831-8921-df0b-43fc-43bad18116d6
ms.date: 06/08/2017
---


# DefaultWebOptions.BrowserLevel Property (Word)

Returns or sets a  **WdBrowserLevel** constant that represents the level of the Web browser for which you want to target new Web pages created in Microsoft Word. Read/write.


## Syntax

 _expression_ . **BrowserLevel**

 _expression_ Required. A variable that represents a **[DefaultWebOptions](defaultweboptions-object-word.md)** collection.


## Remarks

After you set the  **BrowserLevel** property on the **DefaultWebOptions** object, the **BrowserLevel** property of any new Web pages you create in Word will be the same as the global setting.


## Example

This example sets Word to optimize new Web pages for Microsoft Internet Explorer 5 and creates a Web page based on this setting.


```vb
With Application.DefaultWebOptions 
 .BrowserLevel = wdBrowserLevelMicrosoftInternetExplorer5 
 .OptimizeForBrowser = True 
End With 
Documents.Add DocumentType:=wdNewWebPage
```


## See also


#### Concepts


[DefaultWebOptions Object](defaultweboptions-object-word.md)

