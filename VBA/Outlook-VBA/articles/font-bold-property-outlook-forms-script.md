---
title: Font.Bold Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: b1b11748-53e7-0bcd-5c5e-3ad4d4b232b0
ms.date: 06/08/2017
---


# Font.Bold Property (Outlook Forms Script)

Returns or sets a  **Boolean** that specifies whether text is normal or bold. Read/write.


## Syntax

 _expression_. **Bold**

 _expression_A variable that represents a  **Font** object.


## Remarks

 **True** to indicate that the text with this font is bold, **False** otherwise.

The font's appearance on screen and in print may differ, depending on your computer and printer. If you select a font that your system can't display with the specified attribute or that isn't installed, Windows substitutes a similar font. The substitute font will be as similar as possible to the font originally requested.

Changing the value of  **Bold** also changes the value of **[Weight](font-weight-property-outlook-forms-script.md)**. Setting  **Bold** to **True** sets **Weight** to 700; setting **Bold** to **False** sets **Weight** to 400. Conversely, setting **Weight** to anything over 550 sets **Bold** to **True**; setting  **Weight** to 550 or less sets **Bold** to **False**.


