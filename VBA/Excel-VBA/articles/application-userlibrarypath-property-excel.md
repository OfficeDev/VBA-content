---
title: Application.UserLibraryPath Property (Excel)
keywords: vbaxl10.chm133249
f1_keywords:
- vbaxl10.chm133249
ms.prod: excel
api_name:
- Excel.Application.UserLibraryPath
ms.assetid: 48e66da8-4db9-1262-9c0b-3a7f9f8e43ae
ms.date: 06/08/2017
---


# Application.UserLibraryPath Property (Excel)

Returns the path to the location on the user's computer where the COM add-ins are installed. Read-only  **String** .


## Syntax

 _expression_ . **UserLibraryPath**

 _expression_ A variable that represents an **Application** object.


## Example

This example determines where the COM add-ins are installed on the user's computer and assigns the string to the variable  `strLibPath`.


```
strLibPath = Application.UserLibraryPath
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

