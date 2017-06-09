---
title: OldHeight, OldWidth Properties
keywords: fm20.chm2001620
f1_keywords:
- fm20.chm2001620
ms.prod: office
ms.assetid: cd2c0dfb-85f3-2381-128b-4d964829e7b0
ms.date: 06/08/2017
---


# OldHeight, OldWidth Properties



Returns the previous height or width, in [points](vbe-glossary.md), of the control.
 **Syntax**
 _object_. **OldHeight**
 _object_. **OldWidth**
The  **OldHeight** and **OldWidth** property syntaxes have these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
 **Remarks**
 **OldHeight** and **OldWidth** are read-only.
The  **OldHeight** and **OldWidth** properties are automatically updated when you move or size a control. If you change the size of a control, the **Height** and **Width** properties store the new height and **OldHeight** and **OldWidth** store the previous height.
These properties are valid only in the Layout event.

