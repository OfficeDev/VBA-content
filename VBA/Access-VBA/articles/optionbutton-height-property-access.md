---
title: OptionButton.Height Property (Access)
keywords: vbaac10.chm10585
f1_keywords:
- vbaac10.chm10585
ms.prod: access
api_name:
- Access.OptionButton.Height
ms.assetid: d3a95041-1e8f-5a02-019e-ecdb2f795bf0
ms.date: 06/08/2017
---


# OptionButton.Height Property (Access)

Gets or sets the height of the specified object in twips. Read/write  **Integer**.


## Syntax

 _expression_. **Height**

 _expression_ A variable that represents an **OptionButton** object.


## Remarks

For report controls, you can set the  **Height** property when you print or preview a report only by using a macro or an event procedure specified in a section's **OnFormat** event property setting.

Microsoft Access automatically sets the  **Height** property when you create or size a control or when you size a window in form Design View or report Design view.

The height of controls is measured from the center of their borders so controls with different border widths align correctly. 


## See also


#### Concepts


[OptionButton Object](optionbutton-object-access.md)

