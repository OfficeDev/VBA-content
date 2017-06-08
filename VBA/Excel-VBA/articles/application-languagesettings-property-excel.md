---
title: Application.LanguageSettings Property (Excel)
keywords: vbaxl10.chm133251
f1_keywords:
- vbaxl10.chm133251
ms.prod: excel
api_name:
- Excel.Application.LanguageSettings
ms.assetid: 631879d9-f43f-4985-32d0-77bf314956eb
ms.date: 06/08/2017
---


# Application.LanguageSettings Property (Excel)

Returns the  **[LanguageSettings](http://msdn.microsoft.com/library/936f7d61-87e5-e153-08d4-f8c5c8ef0710%28Office.15%29.aspx)** object, which contains information about the language settings in Microsoft Excel. Read-only.


## Syntax

 _expression_ . **LanguageSettings**

 _expression_ A variable that represents an **Application** object.


## Example

This example returns the language identifier for the language you selected when you installed Microsoft Excel.


```vb
Set objLangSet = Application.LanguageSettings 
MsgBox objLangSet.LanguageID(msoLanguageIDInstall)
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

