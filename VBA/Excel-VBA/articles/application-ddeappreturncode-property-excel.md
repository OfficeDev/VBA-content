---
title: Application.DDEAppReturnCode Property (Excel)
keywords: vbaxl10.chm183089
f1_keywords:
- vbaxl10.chm183089
ms.prod: excel
api_name:
- Excel.Application.DDEAppReturnCode
ms.assetid: 9b55dcce-eea8-a8b7-dace-296191de18a4
ms.date: 06/08/2017
---


# Application.DDEAppReturnCode Property (Excel)

Returns the application-specific DDE return code that was contained in the last DDE acknowledge message received by Microsoft Excel. Read-only  **Long** .


## Syntax

 _expression_ . **DDEAppReturnCode**

 _expression_ A variable that represents an **Application** object.


## Example

This example sets the variable  `appErrorCode` to the DDE return code.


```
appErrorCode = Application.DDEAppReturnCode
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

