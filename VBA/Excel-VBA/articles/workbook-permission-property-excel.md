---
title: Workbook.Permission Property (Excel)
keywords: vbaxl10.chm199220
f1_keywords:
- vbaxl10.chm199220
ms.prod: excel
api_name:
- Excel.Workbook.Permission
ms.assetid: ef04f56e-a04d-c3d9-fdda-611be7bf9d39
ms.date: 06/08/2017
---


# Workbook.Permission Property (Excel)

Returns a  **Permission** object that represents the permission settings in the specified workbook.


## Syntax

 _expression_ . **Permission**

 _expression_ A variable that represents a **Workbook** object.


## Example

The following example returns the permission settings for the active workbook.


```vb
Dim objPermission As Permission 
 
Set objPermission = ActiveWorkbook.Permission
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

