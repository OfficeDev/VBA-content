---
title: Phonetics.Add Method (Excel)
keywords: vbaxl10.chm658082
f1_keywords:
- vbaxl10.chm658082
ms.prod: excel
api_name:
- Excel.Phonetics.Add
ms.assetid: 2a60a1cd-e15e-1341-2de8-953aa999ac07
ms.date: 06/08/2017
---


# Phonetics.Add Method (Excel)

Adds phonetic text to the specified cell.


## Syntax

 _expression_ . **Add**( **_Start_** , **_Length_** , **_Text_** )

 _expression_ A variable that represents a **Phonetics** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Start_|Required| **Long**|The position that represents the first character in the specified cell.|
| _Length_|Required| **Long**|The number of characters from the  **Start** position to the end of the text in the cell.|
| _Text_|Required| **String**|Collectively, the characters that represent the phonetic text in the cell.|

## Example

This example adds three phonetic text strings to the active cell. The example then sets the character type to Hiragana, sets the font color to blue, and sets the text to visible.


```vb
ActiveCell.FormulaR1C1 = "?????????" 
ActiveCell.Phonetics.Add Start:=1, Length:=3, Text:="??????" 
ActiveCell.Phonetics.Add Start:=4, Length:=3, Text:="????" 
ActiveCell.Phonetics.CharacterType = xlHiragana 
ActiveCell.Phonetics.Font.Color = vbBlue 
ActiveCell.Phonetics.Visible = True
```


## See also


#### Concepts


[Phonetics Object](phonetics-object-excel.md)

