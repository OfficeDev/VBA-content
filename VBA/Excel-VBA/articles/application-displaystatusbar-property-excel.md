---
title: Application.DisplayStatusBar Property (Excel)
keywords: vbaxl10.chm133127
f1_keywords:
- vbaxl10.chm133127
ms.prod: excel
api_name:
- Excel.Application.DisplayStatusBar
ms.assetid: bf70a679-bd50-cce7-0dc0-0dc57835038c
ms.date: 06/08/2017
---


# Application.DisplayStatusBar Property (Excel)

 **True** if the status bar is displayed. Read/write **Boolean** .


## Syntax

 _expression_ . **DisplayStatusBar**

 _expression_ A variable that represents an **Application** object.


## Example

This example saves the current state of the  **DisplayStatusBar** property and then sets the property to **True** so that the status bar is visible.


```vb
saveStatusBar = Application.DisplayStatusBar 
Application.DisplayStatusBar = True
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

