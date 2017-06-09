---
title: Application.ActivateMicrosoftApp Method (Excel)
keywords: vbaxl10.chm133074
f1_keywords:
- vbaxl10.chm133074
ms.prod: excel
api_name:
- Excel.Application.ActivateMicrosoftApp
ms.assetid: e11d8165-5aad-2b1d-f9d1-797038d96afb
ms.date: 06/08/2017
---


# Application.ActivateMicrosoftApp Method (Excel)

Activates a Microsoft application. If the application is already running, this method activates the running application. If the application isn't running, this method starts a new instance of the application.


## Syntax

 _expression_ . **ActivateMicrosoftApp**( **_Index_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **[XlMSApplication](xlmsapplication-enumeration-excel.md)**|Specifies the Microsoft application to activate.|

## Example

This example starts and activates Word.


```vb
Application.ActivateMicrosoftApp xlMicrosoftWord
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

