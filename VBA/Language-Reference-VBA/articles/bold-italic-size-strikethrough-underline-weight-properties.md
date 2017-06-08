---
title: Bold, Italic, Size, StrikeThrough, Underline, Weight Properties
keywords: fm20.chm5225008
f1_keywords:
- fm20.chm5225008
ms.prod: office
ms.assetid: 1bac5191-4c72-8942-e56f-94cc87647f0f
ms.date: 06/08/2017
---


# Bold, Italic, Size, StrikeThrough, Underline, Weight Properties



Specifies the visual attributes of text on a displayed or printed form.
 **Syntax**
 _object_. **Bold** [= _Boolean_ ]
 _object_. **Italic** [= _Boolean_ ]
 _object_. **Size** [= _Currency_ ]
 _object_. **StrikeThrough** [= _Boolean_ ]
 _object_. **Underline** [= _Boolean_ ]
 _object_. **Weight** [= _Integer_ ]
The  **Bold**, **Italic**, **Size**, **StrikeThrough**, **Underline**, and **Weight** property syntaxes have these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object name.|
| _Boolean_|Optional. Specifies the font style.|
| _Currency_|Optional. A number indicating the font size.|
| _Integer_|Optional. Specifies the font style.|
The settings for  _Boolean_ are:


|**Value**|**Description**|
|:-----|:-----|
|**True**|The text has the specified attribute (that is bold, italic, size, strikethrough or underline marks, or weight).|
|**False**|The text does not have the specified attribute (default).|
The  **Weight** property accepts values from 0 to 1000. A value of zero allows the system to pick the most appropriate weight. A value from 1 to 1000 indicates a specific weight, where 1 represents the lightest type and 1000 represents the darkest type.
 **Remarks**
These properties define the visual characteristics of text. The  **Bold** property determines whether text is normal or bold. The **Italic** property determines whether text is normal or italic. The **Size** property determines the height, in[points](vbe-glossary.md), of displayed text. The  **Underline** property determines whether text is underlined. The **StrikeThrough** property determines whether the text appears with strikethrough marks. The **Weight** property determines the darkness of the type.
The font's appearance on screen and in print may differ, depending on your computer and printer. If you select a font that your system can't display with the specified attribute or that isn't installed, the operating system substitutes a similar font. The substitute font will be as similar as possible to the font originally requested.
Changing the value of  **Bold** also changes the value of **Weight**. Setting **Bold** to **True** sets **Weight** to 700; setting **Bold** to **False** sets **Weight** to 400. Conversely, setting **Weight** to anything over 550 sets **Bold** to **True**; setting **Weight** to 550 or less sets **Bold** to **False**.
The default point size is determined by the operating system.

