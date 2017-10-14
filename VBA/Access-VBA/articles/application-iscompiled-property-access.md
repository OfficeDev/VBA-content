---
title: Application.IsCompiled Property (Access)
keywords: vbaac10.chm12567
f1_keywords:
- vbaac10.chm12567
ms.prod: access
api_name:
- Access.Application.IsCompiled
ms.assetid: c3b80c32-2aba-432c-1909-4c8172a3bebf
ms.date: 06/08/2017
---


# Application.IsCompiled Property (Access)

The  **IsCompiled** property returns a **Boolean** value indicating whether the Visual Basic project is in a compiled state. Read-only **Boolean**.


## Syntax

 _expression_. **IsCompiled**

 _expression_ A variable that represents an **Application** object.


## Remarks

The  **IsCompiled** property of the **[Application](application-object-access.md)** object is **False** when the project has never been fully compiled, if a module has been added, edited, or deleted after compilation, or if a module hasn't been saved in a compiled state.


## See also


#### Concepts


[Application Object](application-object-access.md)

