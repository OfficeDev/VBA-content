---
title: Image.BackStyle Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 1058cd2e-936e-17d5-9276-2a7130ebc3ef
ms.date: 06/08/2017
---


# Image.BackStyle Property (Outlook Forms Script)

Returns or sets an  **Integer** that specifies the background style for an object. Read/write.


## Syntax

 _expression_. **BackStyle**

 _expression_A variable that represents an  **Image** object.


## Remarks

The possible values are 0 and 1. 0 represents the background as transparent, 1 represents the background as opaque.

The  **BackStyle** property determines whether a control is transparent. If **BackStyle** is 1, the control is not transparent and you cannot see anything behind the control on a form. If **BackStyle** is 0, you can see through the control and look at anything on the form located behind the control. The **[BackColor](image-backcolor-property-outlook-forms-script.md)** property is only valid if the **BackStyle** property is set to 1.

 **BackStyle** does not affect the transparency of bitmaps. You must use a picture editor such as Paintbrush to make a bitmap transparent. Not all controls support transparent bitmaps.


