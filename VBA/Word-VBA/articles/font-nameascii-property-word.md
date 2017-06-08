---
title: Font.NameAscii Property (Word)
keywords: vbawd10.chm156369053
f1_keywords:
- vbawd10.chm156369053
ms.prod: word
api_name:
- Word.Font.NameAscii
ms.assetid: 9725a12b-0dd2-0bf7-faa6-2c2b68107771
ms.date: 06/08/2017
---


# Font.NameAscii Property (Word)

Returns or sets the font used for Latin text (characters with character codes from 0 (zero) through 127). Read/write  **String** .


## Syntax

 _expression_ . **NameAscii**

 _expression_ An expression that returns a **[Font](font-object-word.md)** object.


## Remarks

In the U.S. English version of Microsoft Word, the default value of this property is Times New Roman. Use the  **[Name](font-name-property-word.md)** property to change the font that's applied to all text and that appears in the **Font** box on the **Formatting** toolbar.


## Example

This example sets the font used for Latin text.


```
Selection.Font.NameAscii = "Century"
```


## See also


#### Concepts


[Font Object](font-object-word.md)

