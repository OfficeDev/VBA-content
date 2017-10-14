---
title: Phonetics.CharacterType Property (Excel)
keywords: vbaxl10.chm658077
f1_keywords:
- vbaxl10.chm658077
ms.prod: excel
api_name:
- Excel.Phonetics.CharacterType
ms.assetid: b61c3bd5-86dc-baed-e47f-62d522fca290
ms.date: 06/08/2017
---


# Phonetics.CharacterType Property (Excel)

Returns or sets the type of phonetic text in the specified cell. Read/write  **[XlPhoneticCharacterType](xlphoneticcharactertype-enumeration-excel.md)** .


## Syntax

 _expression_ . **CharacterType**

 _expression_ A variable that represents a **Phonetics** object.


## Example

This example changes the first phonetic text string in the active cell from Furigana to Hiragana.


```vb
ActiveCell.Phonetics(1).CharacterType = xlHiragana
```


## See also


#### Concepts


[Phonetics Object](phonetics-object-excel.md)

