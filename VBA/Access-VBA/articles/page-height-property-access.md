---
title: Page.Height Property (Access)
keywords: vbaac10.chm12158
f1_keywords:
- vbaac10.chm12158
ms.prod: access
api_name:
- Access.Page.Height
ms.assetid: df6c7cc3-bcf5-6607-144a-383a1f26d21e
ms.date: 06/08/2017
---


# Page.Height Property (Access)

Gets or sets the height of the specified object in twips. Read/write  **Integer**.


## Syntax

 _expression_. **Height**

 _expression_ A variable that represents a **Page** object.


## Remarks

For report controls, you can set the  **Height** property when you print or preview a report only by using a macro or an event procedure specified in a section's **OnFormat** event property setting.

Microsoft Access automatically sets the  **Height** property when you create or size a control or when you size a window in form Design View or report Design view.

The height of controls is measured from the center of their borders so controls with different border widths align correctly. 


## See also


#### Concepts


[Page Object](page-object-access.md)

