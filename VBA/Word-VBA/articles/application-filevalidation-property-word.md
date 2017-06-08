---
title: Application.FileValidation Property (Word)
keywords: vbawd10.chm158335469
f1_keywords:
- vbawd10.chm158335469
ms.prod: word
api_name:
- Word.Application.FileValidation
ms.assetid: 2f88d1a7-98a7-9ec6-09b3-a09c1a934e01
ms.date: 06/08/2017
---


# Application.FileValidation Property (Word)

Returns or sets how Word will validate files before opening them. Read/write [MsoFileValidationMode](http://msdn.microsoft.com/library/2501a3a5-2053-9fc6-7a3f-bca2fe27f6d1%28Office.15%29.aspx).


## Syntax

 _expression_ . **FileValidation**

 _expression_ A variable that represents an **[Application](application-object-word.md)** object.


### Return Value

[MsoFileValidationMode](http://msdn.microsoft.com/library/2501a3a5-2053-9fc6-7a3f-bca2fe27f6d1%28Office.15%29.aspx)


## Remarks

Files that do not pass validation will be opened in a [Protected View window](protectedviewwindow-object-word.md). The  **FileValidation** property is per session only. If you set the **FileValidation** property, that setting will remain in effect for the entire session the application is open.


## See also


#### Concepts


[Application Object](application-object-word.md)

