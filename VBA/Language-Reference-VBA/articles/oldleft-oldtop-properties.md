---
title: OldLeft, OldTop Properties
keywords: fm20.chm2001630
f1_keywords:
- fm20.chm2001630
ms.prod: office
ms.assetid: 034354a8-6a04-a3cc-c28a-3af3cdf2ed65
ms.date: 06/08/2017
---


# OldLeft, OldTop Properties



Returns the distance, in [points](vbe-glossary.md), between the previous position of a control and the left or top edge of the form that contains it.
 **Syntax**
 _object_. **OldLeft**
 _object_. **OldTop**
The  **OldLeft** and **OldTop** property syntaxes have these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
 **Remarks**
 **OldLeft** and **OldTop** are read-only.
The  **OldLeft** and **OldTop** properties are automatically updated when you move or size a control. If you move a control, the **Left** and **Top** properties store the new distance from the control to the left edge of its[container](vbe-glossary.md) and **OldLeft** and **OldTop** store the previous value of **Left**.
 **OldLeft** and **OldTop** are valid only in the Layout event.

