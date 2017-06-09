---
title: CustomControl.Height Property (Access)
keywords: vbaac10.chm12023
f1_keywords:
- vbaac10.chm12023
ms.prod: access
api_name:
- Access.CustomControl.Height
ms.assetid: 1e482282-30f7-139e-dd89-40cf89139a2e
ms.date: 06/08/2017
---


# CustomControl.Height Property (Access)

Gets or sets the height of the specified object in twips. Read/write  **Integer**.


## Syntax

 _expression_. **Height**

 _expression_ A variable that represents a **CustomControl** object.


## Remarks

For report controls, you can set the  **Height** property when you print or preview a report only by using a macro or an event procedure specified in a section's **OnFormat** event property setting.

Microsoft Access automatically sets the  **Height** property when you create or size a control or when you size a window in form Design View or report Design view.

The height of controls is measured from the center of their borders so controls with different border widths align correctly. 


## See also


#### Concepts


[CustomControl Object](customcontrol-object-access.md)

