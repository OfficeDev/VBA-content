---
title: SharingItem.BodyFormat Property (Outlook)
keywords: vbaol11.chm675
f1_keywords:
- vbaol11.chm675
ms.prod: outlook
api_name:
- Outlook.SharingItem.BodyFormat
ms.assetid: 60a18df9-8882-a5a2-efb9-cc59206f7345
ms.date: 06/08/2017
---


# SharingItem.BodyFormat Property (Outlook)

Returns or sets an  **[OlBodyFormat](olbodyformat-enumeration-outlook.md)** constant indicating the format of the body text. Read/write.


## Syntax

 _expression_ . **BodyFormat**

 _expression_ A variable that represents a **SharingItem** object.


## Remarks

The body text format determines the standard used to display the text of the message. Microsoft Outlook provides three body text format options: Plain Text, Rich Text (RTF), and HTML.

All text formatting will be lost when the  **BodyFormat** property is switched from RTF to HTML and vice-versa.


## See also


#### Concepts


[SharingItem Object](sharingitem-object-outlook.md)

