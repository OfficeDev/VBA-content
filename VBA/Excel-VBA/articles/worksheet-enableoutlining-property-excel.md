---
title: Worksheet.EnableOutlining Property (Excel)
keywords: vbaxl10.chm175096
f1_keywords:
- vbaxl10.chm175096
ms.prod: excel
api_name:
- Excel.Worksheet.EnableOutlining
ms.assetid: db849ddf-871d-19cd-9765-3194a8c1e34e
ms.date: 06/08/2017
---


# Worksheet.EnableOutlining Property (Excel)

 **True** if outlining symbols are enabled when user-interface-only protection is turned on. Read/write **Boolean** .


## Syntax

 _expression_ . **EnableOutlining**

 _expression_ A variable that represents a **Worksheet** object.


## Remarks

This example enables outlining symbols on a protected worksheet.


## Example


```vb
ActiveSheet.EnableOutlining = True 
ActiveSheet.Protect contents:=True, userInterfaceOnly:=True
```


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

