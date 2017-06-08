---
title: Application.FileValidationPivot Property (Excel)
keywords: vbaxl10.chm133336
f1_keywords:
- vbaxl10.chm133336
ms.prod: excel
api_name:
- Excel.Application.FileValidationPivot
ms.assetid: 3cf6e177-9dbe-8ee8-3d84-599d7e2221da
ms.date: 06/08/2017
---


# Application.FileValidationPivot Property (Excel)

Returns or sets how Excel will validate the contents of the data caches for PivotTable reports. Read/write


## Syntax

 _expression_ . **FileValidationPivot**

 _expression_ A variable that represents an **[Application](application-object-excel.md)** object.


### Return Value

 **[XlFileValidationPivotMode](xlfilevalidationpivotmode-enumeration-excel.md)**


## Remarks

Files that contain data caches that do not validate will be opened in a  **Protected View** window. If you set the **FileValidationPivot** property, that setting will remain in effect for the entire session the application is open.


## See also


#### Concepts


[Application Object](application-object-excel.md)

