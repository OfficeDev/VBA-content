---
title: TabStrip.ClientLeft Property (Outlook Forms Script)
keywords: olfm10.chm2000900
f1_keywords:
- olfm10.chm2000900
ms.prod: outlook
ms.assetid: 4774cba6-430d-da76-f67f-fede5aec6eea
ms.date: 06/08/2017
---


# TabStrip.ClientLeft Property (Outlook Forms Script)

Returns a  **Single** value that represents the location of the left edge of the display area of a **[TabStrip](tabstrip-object-outlook-forms-script.md)**. Read-only.


## Syntax

 _expression_. **ClientLeft**

 _expression_A variable that represents a  **TabStrip** object.


## Remarks

For  **[ClientHeight](tabstrip-clientheight-property-outlook-forms-script.md)** and **[ClientWidth](tabstrip-clientwidth-property-outlook-forms-script.md)**, specifies the distance, in points, from respectively the top and left edge of the TabStrip's container. For  **ClientLeft** and **[ClientTop](tabstrip-clienttop-property-outlook-forms-script.md)**, specifies the location, in points, of respectively the top and left edges of the TabStrip's container.

At run time,  **ClientLeft**,  **ClientTop**,  **ClientHeight**, and  **ClientWidth** automatically store the coordinates and dimensions of the **TabStrip** control's internal area, which is shared by objects in the **TabStrip**.


