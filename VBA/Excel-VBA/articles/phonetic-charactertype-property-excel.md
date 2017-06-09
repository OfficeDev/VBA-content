---
title: Phonetic.CharacterType Property (Excel)
keywords: vbaxl10.chm628074
f1_keywords:
- vbaxl10.chm628074
ms.prod: excel
api_name:
- Excel.Phonetic.CharacterType
ms.assetid: 2c8ba9b0-1d87-7627-7083-31c9260b68b5
ms.date: 06/08/2017
---


# Phonetic.CharacterType Property (Excel)

Returns or sets the type of phonetic text in the specified cell. Read/write  **[XlPhoneticCharacterType](xlphoneticcharactertype-enumeration-excel.md)** .


## Syntax

 _expression_ . **CharacterType**

 _expression_ A variable that represents a **Phonetic** object.


## Example

This example changes the first phonetic text string in the active cell from Furigana to Hiragana.


```vb
ActiveCell.Phonetics(1).CharacterType = xlHiragana
```


## See also


#### Concepts


[Phonetic Object](phonetic-object-excel.md)

