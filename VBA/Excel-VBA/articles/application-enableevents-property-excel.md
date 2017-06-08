---
title: Application.EnableEvents Property (Excel)
keywords: vbaxl10.chm133240
f1_keywords:
- vbaxl10.chm133240
ms.prod: excel
api_name:
- Excel.Application.EnableEvents
ms.assetid: 5e14ce7b-02f6-03d4-2dfc-1df05a032301
ms.date: 06/08/2017
---


# Application.EnableEvents Property (Excel)

 **True** if events are enabled for the specified object. Read/write **Boolean** .


## Syntax

 _expression_ . **EnableEvents**

 _expression_ A variable that represents an **Application** object.


## Example

This example disables events before a file is saved so that the  **BeforeSave** event doesn't occur.


```vb
Application.EnableEvents = False 
ActiveWorkbook.Save 
Application.EnableEvents = True
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

