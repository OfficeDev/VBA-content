---
title: Language.WritingStyleList Property (Word)
keywords: vbawd10.chm158138386
f1_keywords:
- vbawd10.chm158138386
ms.prod: word
api_name:
- Word.Language.WritingStyleList
ms.assetid: 5a91ecaa-dce0-d9ab-0e25-ec9620fa7119
ms.date: 06/08/2017
---


# Language.WritingStyleList Property (Word)

Returns a string array that contains the names of all writing styles available for the specified language. Read-only  **Variant** .


## Syntax

 _expression_ . **WritingStyleList**

 _expression_ An expression that returns a **[Language](language-object-word.md)** object.


## Example

This example displays each writing style available for U.S. English. Each writing style and its number in the array are also displayed in the Immediate window of the Visual Basic editor.


```vb
Sub WritingStyles() 
 Dim WrStyles As Variant 
 Dim i As Integer 
 
 WrStyles = Languages(wdEnglishUS).WritingStyleList 
 For i = 1 To UBound(WrStyles) 
 MsgBox WrStyles(i) 
 Debug.Print WrStyles(i) &; " [" &; Trim(Str$(i)) &; "]" 
 Next i 
End Sub
```


## See also


#### Concepts


[Language Object](language-object-word.md)

