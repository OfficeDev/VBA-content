---
title: Application.StatusBar Property (Excel)
keywords: vbaxl10.chm133213
f1_keywords:
- vbaxl10.chm133213
ms.prod: excel
api_name:
- Excel.Application.StatusBar
ms.assetid: 91b043d7-b320-da4b-bdc7-3be1e1ffe3c6
ms.date: 06/08/2017
---


# Application.StatusBar Property (Excel)

Returns or sets the text in the status bar. Read/write  **String** .


## Syntax

 _expression_ . **StatusBar**

 _expression_ A variable that represents an **Application** object.


## Remarks

This property returns  **False** if Microsoft Excel has control of the status bar. To restore the default status bar text, set the property to **False** ; this works even if the status bar is hidden.


## Example

This example sets the status bar text to "Please be patient..." before it opens the workbook Large.xls, and then it restores the default text.


```vb
oldStatusBar = Application.DisplayStatusBar 
Application.DisplayStatusBar = True 
Application.StatusBar = "Please be patient..." 
Workbooks.Open filename:="LARGE.XLS" 
Application.StatusBar = False 
Application.DisplayStatusBar = oldStatusBar
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

