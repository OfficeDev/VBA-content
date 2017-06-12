---
title: Font.Weight Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 674d5bd0-3cf7-1330-5d3c-7b742bb3df7c
ms.date: 06/08/2017
---


# Font.Weight Property (Outlook Forms Script)

Returns or sets an  **Integer** that specifies the lthe darkness of the type. Read/write.


## Syntax

 _expression_. **Weight**

 _expression_A variable that represents a  **Font** object.


## Remarks

The  **Weight** property accepts values from 0 to 1000. A value of zero allows the system to pick the most appropriate weight. A value from 1 to 1000 indicates a specific weight, where 1 represents the lightest type and 1000 represents the darkest type.

The font's appearance on screen and in print may differ, depending on your computer and printer. If you select a font that your system can't display with the specified attribute or that isn't installed, Windows substitutes a similar font. The substitute font will be as similar as possible to the font originally requested.

Changing the value of  **[Bold](font-bold-property-outlook-forms-script.md)** also changes the value of **Weight**. Setting  **Bold** to **True** sets **Weight** to 700; setting **Bold** to **False** sets **Weight** to 400. Conversely, setting **Weight** to anything over 550 sets **Bold** to **True**; setting  **Weight** to 550 or less sets **Bold** to **False**.


