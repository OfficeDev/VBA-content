---
title: Phonetic Object (Excel)
keywords: vbaxl10.chm627072
f1_keywords:
- vbaxl10.chm627072
ms.prod: excel
api_name:
- Excel.Phonetic
ms.assetid: 297e85d5-e8f6-6009-c51a-0d3fe01efba0
ms.date: 06/08/2017
---


# Phonetic Object (Excel)

Contains information about a specific phonetic text string in a cell.


## Remarks

 In Microsoft Excel 97, this object contained the formatting attributes for any phonetic text in the specified range.


## Example

Use  **[Phonetics](range-phonetics-property-excel.md)** ( _index_ ), where _index_ is the index number of the phonetic text, to return a single **Phonetic** object. The following example sets the first phonetic text string in the active cell to "フリガナ".


```vb
ActiveCell.Phonetics(1).Text = "フリガナ"
```

The  **[Phonetic](range-phonetic-property-excel.md)** property provides compatibility with earlier versions of Microsoft Excel. You should use **Phonetics** ( _index_ ), where _index_ is the index number of the phonetic text, to return a single **Phonetic** object. To demonstrate compatibility with earlier versions of Microsoft Excel, the following example adds Furigana characters to the range A1:C4. If you add Furigana characters to a range, a new **Phonetic** object is automatically created.




```vb
With Range("A1:C4").Phonetic 
    .CharacterType = xlHiragana 
    .Alignment = xlPhoneticAlignCenter 
    .Font.Name = "MS P ゴシック" 
    .Font.FontStyle = "標準" 
    .Font.Size = 6 
    .Font.Strikethrough = False 
    .Font.Underline = xlUnderlineStyleNone 
    .Font.ColorIndex = xlAutomatic 
    .Visible = True 
End With
```


## See also


#### Other resources



[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)

