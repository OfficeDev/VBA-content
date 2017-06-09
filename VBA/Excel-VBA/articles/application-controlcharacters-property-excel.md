---
title: Application.ControlCharacters Property (Excel)
keywords: vbaxl10.chm133238
f1_keywords:
- vbaxl10.chm133238
ms.prod: excel
api_name:
- Excel.Application.ControlCharacters
ms.assetid: 039a266a-e5ae-468e-e3ee-101fa2b12863
ms.date: 06/08/2017
---


# Application.ControlCharacters Property (Excel)

 **True** if Microsoft Excel displays control characters for right-to-left languages. Read/write **Boolean** .


## Syntax

 _expression_ . **ControlCharacters**

 _expression_ A variable that represents an **Application** object.


## Remarks

This property can be set only when right to left language support has been installed and selected.


## Example

This example sets Microsoft Excel to interpret control characters.


```vb
Application.ControlCharacters = True
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

