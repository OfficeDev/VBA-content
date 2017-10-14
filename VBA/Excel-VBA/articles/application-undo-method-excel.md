---
title: Application.Undo Method (Excel)
keywords: vbaxl10.chm133221
f1_keywords:
- vbaxl10.chm133221
ms.prod: excel
api_name:
- Excel.Application.Undo
ms.assetid: b56bb8a0-2cd1-356a-03ba-47eb6f56f455
ms.date: 06/08/2017
---


# Application.Undo Method (Excel)

Cancels the last user-interface action.


## Syntax

 _expression_ . **Undo**

 _expression_ A variable that represents an **Application** object.


## Remarks

This method undoes only the last action taken by the user before running the macro, and it must be the first line in the macro. It cannot be used to undo Visual Basic commands.


## Example

This example cancels the last user-interface action. The example must be the first line in a macro.


```vb
Application.Undo
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

