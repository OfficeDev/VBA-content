---
title: Application.FileValidation Property (PowerPoint)
keywords: vbapp10.chm502069
f1_keywords:
- vbapp10.chm502069
ms.prod: powerpoint
api_name:
- PowerPoint.Application.FileValidation
ms.assetid: 90cc8bff-df3b-7a57-adcc-bbfb9c677468
ms.date: 06/08/2017
---


# Application.FileValidation Property (PowerPoint)

Returns or sets a value that indicates how PowerPoint will validate files before opening them. Read/write


## Syntax

 _expression_. **FileValidation**

 _expression_ A variable that represents an **Application** object.


### Return Value

 **[MsoFileValidationMode](http://msdn.microsoft.com/library/2501a3a5-2053-9fc6-7a3f-bca2fe27f6d1%28Office.15%29.aspx)**


## Remarks

Files that do not pass validation will be opened in a  **Protected View** window. If you set the **FileValidation** property, that setting will remain in effect for the entire session during which the application is open.


## See also


#### Concepts


[Application Object](application-object-powerpoint.md)

