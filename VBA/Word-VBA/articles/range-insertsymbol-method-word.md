---
title: Range.InsertSymbol Method (Word)
keywords: vbawd10.chm157155492
f1_keywords:
- vbawd10.chm157155492
ms.prod: word
api_name:
- Word.Range.InsertSymbol
ms.assetid: 2fa843fa-4966-a4e6-1411-028b14029bdf
ms.date: 06/08/2017
---


# Range.InsertSymbol Method (Word)

Inserts a symbol in place of the specified range.


## Syntax

 _expression_ . **InsertSymbol**( **_CharacterNumber_** , **_Font_** , **_Unicode_** , **_Bias_** )

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _CharacterNumber_|Required| **Long**|The character number for the specified symbol. This value will always be the sum of 31 and the number that corresponds to the position of the symbol in the table of symbols (counting from left to right). For example, to specify a delta character at position 37 in the table of symbols in the Symbol font, set CharacterNumber to 68.|
| _Font_|Optional| **Variant**|The name of the font that contains the symbol.|
| _Unicode_|Optional| **Variant**| **True** to insert the unicode character specified by CharacterNumber; **False** to insert the ANSI character specified by CharacterNumber. The default value is **False** .|
| _Bias_|Optional| **Variant**|Sets the font bias for symbols. This argument is useful for setting the correct font bias for East Asian characters. Can be one of the  **WdFontBias** constants. This argument may not be available to you, depending on the language support (U.S. English, for example) that you?ve selected or installed.|

## Remarks

If you don't want to replace the range, use the  **Collapse** method before you use this method.


## Example

This example inserts a double-headed arrow at the end of the document.


```vb
With ActiveDocument.Range 
 .Collapse Direction:=wdCollapseStart 
 .InsertSymbol CharacterNumber:=171, _ 
 Font:="Symbol", Unicode:=False 
End With
```


## See also


#### Concepts


[Range Object](range-object-word.md)

