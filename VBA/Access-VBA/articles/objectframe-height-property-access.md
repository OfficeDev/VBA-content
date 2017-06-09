---
title: ObjectFrame.Height Property (Access)
keywords: vbaac10.chm11591
f1_keywords:
- vbaac10.chm11591
ms.prod: access
api_name:
- Access.ObjectFrame.Height
ms.assetid: 53d6085e-e01e-5260-0802-3958f62e378a
ms.date: 06/08/2017
---


# ObjectFrame.Height Property (Access)

Gets or sets the height of the specified object in twips. Read/write  **Integer**.


## Syntax

 _expression_. **Height**

 _expression_ A variable that represents an **ObjectFrame** object.


## Remarks

For report controls, you can set the  **Height** property when you print or preview a report only by using a macro or an event procedure specified in a section's **OnFormat** event property setting.

Microsoft Access automatically sets the  **Height** property when you create or size a control or when you size a window in form Design View or report Design view.

The height of controls is measured from the center of their borders so controls with different border widths align correctly. 


## See also


#### Concepts


[ObjectFrame Object](objectframe-object-access.md)

