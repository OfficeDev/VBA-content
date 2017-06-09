---
title: Application.TransitionMenuKeyAction Property (Excel)
keywords: vbaxl10.chm133219
f1_keywords:
- vbaxl10.chm133219
ms.prod: excel
api_name:
- Excel.Application.TransitionMenuKeyAction
ms.assetid: 8f278d3b-9902-597a-9e4d-7f2fc3f22469
ms.date: 06/08/2017
---


# Application.TransitionMenuKeyAction Property (Excel)

Returns or sets the action taken when the Microsoft Excel menu key is pressed. Can be either  **xlExcelMenus** or **xlLotusHelp** . Read/write **Long** .


## Syntax

 _expression_ . **TransitionMenuKeyAction**

 _expression_ A variable that represents an **Application** object.


## Example

This example sets the Microsoft Excel menu key to run Lotus 1-2-3 Help when it is pressed.


```vb
Application.TransitionMenuKeyAction = xlLotusHelp 

```


## See also


#### Concepts


[Application Object](application-object-excel.md)

