---
title: Application.Version Property (Excel)
keywords: vbaxl10.chm133228
f1_keywords:
- vbaxl10.chm133228
ms.prod: excel
api_name:
- Excel.Application.Version
ms.assetid: 071cad0c-1cc0-8972-76f8-7c04d42765bd
ms.date: 06/08/2017
---


# Application.Version Property (Excel)

Returns a  **String** value that represents the Microsoft Excel version number.


## Syntax

 _expression_ . **Version**

 _expression_ A variable that represents an **Application** object.


## Example

This example displays a message box that contains the Microsoft Excel version number and the name of the operating system.


```vb
MsgBox "Welcome to Microsoft Excel version " &; _ 
 Application.Version &; " running on " &; _ 
 Application.OperatingSystem &; "!"
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

