---
title: Image.Height Property (Access)
keywords: vbaac10.chm10379
f1_keywords:
- vbaac10.chm10379
ms.prod: access
api_name:
- Access.Image.Height
ms.assetid: cfc48a43-736f-58c0-95a5-fe248de8b9d3
ms.date: 06/08/2017
---


# Image.Height Property (Access)

Gets or sets the height of the specified object in twips. Read/write  **Integer**.


## Syntax

 _expression_. **Height**

 _expression_ A variable that represents an **Image** object.


### Return Value

Integer


## Remarks

For report controls, you can set the  **Height** property when you print or preview a report only by using a macro or an event procedure specified in a section's **OnFormat** event property setting.

Microsoft Access automatically sets the  **Height** property when you create or size a control or when you size a window in form Design View or report Design view.

The height of controls is measured from the center of their borders so controls with different border widths align correctly.


## See also


#### Concepts


[Image Object](image-object-access.md)

