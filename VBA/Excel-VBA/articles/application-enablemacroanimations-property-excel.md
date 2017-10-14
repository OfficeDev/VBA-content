---
title: Application.EnableMacroAnimations Property (Excel)
keywords: vbaxl10.chm133340
f1_keywords:
- vbaxl10.chm133340
ms.prod: excel
ms.assetid: b1befccc-4f27-862b-8ab3-c862b5cb79b3
ms.date: 06/08/2017
---


# Application.EnableMacroAnimations Property (Excel)

Controls whether macro animations are enabled.  **True** if user interface animations or chart animations are enabled. Is set to **False** (no animation) by default. If it is set to **True** during the running of a macro, it will enable animation and then will reset to **False** after the macro runs. Read/write **Boolean** .


## Syntax

 _expression_ . **EnableMacroAnimations**

 _expression_ A variable that represents an **Application** object.


## Example

This example disables animation.


```vb
Application.EnableMacroAnimations = False
```


## Property value

 **BOOL**


## See also


#### Concepts


[Application Object](application-object-excel.md)

