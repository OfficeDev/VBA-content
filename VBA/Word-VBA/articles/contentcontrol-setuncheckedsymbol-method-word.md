---
title: ContentControl.SetUncheckedSymbol Method (Word)
keywords: vbawd10.chm266534942
f1_keywords:
- vbawd10.chm266534942
ms.prod: word
api_name:
- Word.ContentControl.SetUncheckedSymbol
ms.assetid: b563d0cd-e504-7aea-9c91-5c6daf257c35
ms.date: 06/08/2017
---


# ContentControl.SetUncheckedSymbol Method (Word)

Sets the symbol used to represent the unchecked state of a check box content control.


## Syntax

 _expression_ . **SetCheckedSymbol**( **_CharacterNumber_** , **_Font_** )

 _expression_ An expression that returns a **[ContentControl](contentcontrol-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _CharacterNumber_|Required| **Long**|The Unicode character number for the specified symbol. This value will always be the sum of 31 (the number of control symbols at the beginning of the font) and the number that corresponds to the position of the symbol in the table of symbols (counting from left to right). For example, to specify a delta character at position 37 in the table of symbols in the Symbol font, set CharacterNumber to 68.|
| _Font_|Optional| **String**|The name of the font that contains the symbol.|

## Example

The following code example sets the unchecked symbol of the specified content control to the "MS Gothic" font "Ballot Box" symbol.


```vb
Dim objCC As ContentControl 
 
Set objCC = ActiveDocument.ContentControls.Add (wdContentControlCheckbox) 
objCC.SetUncheckedSymbol CharacterNumber:=&;H2610, Font:="MS Gothic" 

```


## See also


#### Concepts


[ContentControl Object](contentcontrol-object-word.md)

