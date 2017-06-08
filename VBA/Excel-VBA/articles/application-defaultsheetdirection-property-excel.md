---
title: Application.DefaultSheetDirection Property (Excel)
keywords: vbaxl10.chm133236
f1_keywords:
- vbaxl10.chm133236
ms.prod: excel
api_name:
- Excel.Application.DefaultSheetDirection
ms.assetid: 33fad777-e2dd-99b5-9b33-a573a729b331
ms.date: 06/08/2017
---


# Application.DefaultSheetDirection Property (Excel)

Returns or sets the default direction in which Microsoft Excel displays new windows and worksheets. Can be one of the following constants:  **xlRTL** (right to left) or **xlLTR** (left to right). Read/write **Long** .


## Syntax

 _expression_ . **DefaultSheetDirection**

 _expression_ A variable that represents an **Application** object.


## Remarks

Some of these constants may not be available to you, depending on the language support (U.S. English, for example) that you've selected or installed.


## Example

This example sets right to left as the default direction.


```vb
Application.DefaultSheetDirection = xlRTL
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

