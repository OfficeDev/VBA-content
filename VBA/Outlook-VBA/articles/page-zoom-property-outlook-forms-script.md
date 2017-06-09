---
title: Page.Zoom Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 8664e1d4-4b2b-5415-f5b4-be11ecde7a17
ms.date: 06/08/2017
---


# Page.Zoom Property (Outlook Forms Script)

Returns or sets an  **Integer** that specifies the percentage to increase or decrease the displayed image. Read/write.


## Syntax

 _expression_. **Zoom**

 _expression_A variable that represents a  **Page** object.


## Remarks

The value of the  **Zoom** property specifies a percentage of image enlargement or reduction by which an image display should change. Values from 10 to 400 are valid. The value specified is a percentage of the object's original size; thus, a setting of 400 means you want to enlarge the image to four times its original size (or 400 percent), while a setting of 10 means you want to reduce the image to one-tenth of its original size (or 10 percent).


