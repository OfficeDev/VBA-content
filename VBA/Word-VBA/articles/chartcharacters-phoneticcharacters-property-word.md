---
title: ChartCharacters.PhoneticCharacters Property (Word)
keywords: vbawd10.chm250742258
f1_keywords:
- vbawd10.chm250742258
ms.prod: word
api_name:
- Word.ChartCharacters.PhoneticCharacters
ms.assetid: 3bf59590-d83c-1d11-f092-61b190cd24ad
ms.date: 06/08/2017
---


# ChartCharacters.PhoneticCharacters Property (Word)

Returns or sets the phonetic text for the object. Read/write  **String** .


## Syntax

 _expression_ . **PhoneticCharacters**

 _expression_ A variable that represents a **[ChartCharacters](chartcharacters-object-word.md)** object.


## Example

The following example replaces the first three characters in the title of the first chart in the active document with Furigana characters.


```vb
With ActiveDocument.InlineShapes(1) 
    If .HasChart Then 
        .Chart.Title.Characters(1,3).PhoneticCharacters = "フリガナ" 
    End If 
End With
```


## See also


#### Concepts


[ChartCharacters Object](chartcharacters-object-word.md)

