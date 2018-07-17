---
title: Application.FileValidation Property (Excel)
keywords: vbaxl10.chm133335
f1_keywords:
- vbaxl10.chm133335
ms.prod: excel
api_name:
- Excel.Application.FileValidation
ms.assetid: 6ec989d0-2ed8-b4d9-997c-4f91507e6fca
ms.date: 06/08/2017
---


# Application.FileValidation Property (Excel)

Returns or sets how Excel will validate files before opening them. Read/write


## Syntax

 _expression_ . **FileValidation**

 _expression_ A variable that represents an **[Application](application-object-excel.md)** object.


### Return Value

 **[MsoFileValidationMode](http://msdn.microsoft.com/library/2501a3a5-2053-9fc6-7a3f-bca2fe27f6d1%28Office.15%29.aspx)**


## Remarks

Files that do not pass validation will be opened in a  **Protected View** window. If you set the **FileValidation** property, that setting will remain in effect for the entire session the application is open.


## See also


#### Concepts


[Application Object](application-object-excel.md)

