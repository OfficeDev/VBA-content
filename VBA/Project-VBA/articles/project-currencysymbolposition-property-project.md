---
title: Project.CurrencySymbolPosition Property (Project)
keywords: vbapj.chm131698
f1_keywords:
- vbapj.chm131698
ms.prod: project-server
api_name:
- Project.Project.CurrencySymbolPosition
ms.assetid: 1ac5a154-370f-53f9-0deb-17ee36ec2ad2
ms.date: 06/08/2017
---


# Project.CurrencySymbolPosition Property (Project)

Gets or sets the location of the currency symbol. Read/write  **PjPlacement**.


## Syntax

 _expression_. **CurrencySymbolPosition**

 _expression_ A variable that represents a **Project** object.


## Remarks

Project sets the  **CurrencySymbolPosition** property equal to the corresponding value in the **Customize Regional Options** dialog box of the Windows Control Panel. The value can be one of the following **[PjPlacement](pjplacement-enumeration-project.md)** constants: **pjBefore**, **pjAfter**, **pjBeforeWithSpace**, or **pjAfterWithSpace**.


