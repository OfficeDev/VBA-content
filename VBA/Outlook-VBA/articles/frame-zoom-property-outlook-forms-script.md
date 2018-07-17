---
title: Frame.Zoom Property (Outlook Forms Script)
keywords: olfm10.chm2002240
f1_keywords:
- olfm10.chm2002240
ms.prod: outlook
ms.assetid: a4f67386-1300-c13c-433c-e60434180a9c
ms.date: 06/08/2017
---


# Frame.Zoom Property (Outlook Forms Script)

Returns or sets an  **Integer** that specifies the percentage to increase or decrease the displayed image. Read/write.


## Syntax

 _expression_. **Zoom**

 _expression_A variable that represents a  **Frame** object.


## Remarks

The value of the  **Zoom** property specifies a percentage of image enlargement or reduction by which an image display should change. Values from 10 to 400 are valid. The value specified is a percentage of the object's original size; thus, a setting of 400 means you want to enlarge the image to four times its original size (or 400 percent), while a setting of 10 means you want to reduce the image to one-tenth of its original size (or 10 percent).


