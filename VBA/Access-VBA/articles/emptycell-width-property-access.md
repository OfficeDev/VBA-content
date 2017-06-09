---
title: EmptyCell.Width Property (Access)
keywords: vbaac10.chm14309
f1_keywords:
- vbaac10.chm14309
ms.prod: access
api_name:
- Access.EmptyCell.Width
ms.assetid: ae5cf65a-4f32-7f34-0202-5f4919b033c7
ms.date: 06/08/2017
---


# EmptyCell.Width Property (Access)

Gets or sets the width of the specified object in twips. Read/write  **Integer**.


## Syntax

 _expression_. **Width**

 _expression_ A variable that represents an **EmptyCell** object.


## Remarks

For report controls, you can set the  **Width** property when you print or preview a report only by using a macro or an event procedure specified in a section's **OnFormat** event property setting.

You can't set this property for an object once the print process has started.

Microsoft Access automatically sets the  **Width** property when you create or size a control or when you size a window in form Design View or report Design view.

The width of forms and reports is measured from the inside of their borders. The width of controls is measured from the center of their borders so controls with different border widths align correctly. The margins for forms and reports are set in the  **Page Setup** dialog box, available by clicking **Page Setup** on the **File** menu.


## See also


#### Concepts


[EmptyCell Object](emptycell-object-access.md)

