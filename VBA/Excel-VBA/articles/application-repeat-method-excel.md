---
title: Application.Repeat Method (Excel)
keywords: vbaxl10.chm133200
f1_keywords:
- vbaxl10.chm133200
ms.prod: excel
api_name:
- Excel.Application.Repeat
ms.assetid: ce8f6340-174e-b6cf-0f99-f39be2cde5c2
ms.date: 06/08/2017
---


# Application.Repeat Method (Excel)

Repeats the last user-interface action.


## Syntax

 _expression_ . **Repeat**

 _expression_ A variable that represents an **Application** object.


## Remarks

This method repeats only the last action taken by the user before running the macro, and it must be the first line in the macro. It cannot be used to repeat Visual Basic commands.


## Example

This example repeats the last user-interface command. The example must be the first line in a macro.


```vb
Application.Repeat
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

