---
title: Section.HasContinued Property (Access)
keywords: vbaac10.chm12212
f1_keywords:
- vbaac10.chm12212
ms.prod: access
api_name:
- Access.Section.HasContinued
ms.assetid: 843cb415-5cab-f380-f6f9-854f91393576
ms.date: 06/08/2017
---


# Section.HasContinued Property (Access)

You can use the  **HasContinued** property to determine if part of the current section begins on the previous page. Read/write **Boolean**.


## Syntax

 _expression_. **HasContinued**

 _expression_ A variable that represents a **Section** object.


## Remarks

The  **HasContinued** property is set by Microsoft Access and is read-only in all views.



|**Value**|**Description**|
|:-----|:-----|
|**True**|Part of the current section has been printed on the previous page.|
|**False**|Part of the current section hasn't been printed on the previous page.|
You can use this property to determine whether to show or hide certain controls depending on the value of the property. For example, you may have a hidden label in a page header containing the text "Continued from previous page". If the value of the  **HasContinued** property is **True**, you can make the hidden label visible.


## See also


#### Concepts


[Section Object](section-object-access.md)

