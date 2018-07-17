---
title: Range.GoTo Method (Word)
keywords: vbawd10.chm157155501
f1_keywords:
- vbawd10.chm157155501
ms.prod: word
api_name:
- Word.Range.GoTo
ms.assetid: 9e7cdfcc-756c-4bc8-902e-12479388ea03
ms.date: 06/08/2017
---


# Range.GoTo Method (Word)

Returns a  **Range** object that represents the start position of the specified item, such as a page, bookmark, or field.


## Syntax

 _expression_ . **GoTo**( **_What_** , **_Which_** , **_Count_** , **_Name_** )

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _What_|Optional| **Variant**|The kind of item to which the range is moved. Can be one of the  **WdGoToItem** constants.|
| _Which_|Optional| **Variant**|The item to which the range is moved. Can be one of the  **WdGoToDirection** constants.|
| _Count_|Optional| **Variant**|The number of the item in the document. The default value is 1. Only positive values are valid. To specify an item that precedes the range, use  **wdGoToPrevious** as the Which argument and specify a Count value.|
| _Name_|Optional| **Variant**|If the What argument is  **wdGoToBookmark** , **wdGoToComment** , **wdGoToField** , or **wdGoToObject** , this argument specifies a name. Only positive values are valid. To specify an item that precedes the range, use **wdGoToPrevious** as the Which argument and specify a Count value.|

## Remarks

The following example moves the range up two lines.


```vb
ActiveDocument.Range.GoTo What:=wdGoToLine, Which:=wdGoToPrevious, Count:=2
```

The following example moves to the next DATE field.




```vb
ActiveDocument.Range.GoTo What:=wdGoToField, Name:="Date"
```

The following example moves the range to the fourth line in the document.




```vb
ActiveDocument.Range.GoTo What:=wdGoToLine, Which:=wdGoToAbsolute, Count:=4
```

The following examples are functionally equivalent; they both move the range to the first heading in the document.




```vb
ActiveDocument.Range.GoTo What:=wdGoToHeading, Which:=wdGoToFirst 
ActiveDocument.Range.GoTo What:=wdGoToHeading, Which:=wdGoToAbsolute, Count:=1
```

When you use the  **GoTo** method with the **wdGoToGrammaticalError** , **wdGoToProofreadingError** , or **wdGoToSpellingError** constant, the **Range** that is returned includes any grammar error text or spelling error text.


## Example

This example moves the insertion point just before the fifth endnote reference mark in the active document.


```vb
If ActiveDocument.Endnotes.Count >= 5 Then 
 ActiveDocument.Range.GoTo What:=wdGoToEndnote, _ 
 Which:=wdGoToAbsolute, Count:=5 
End If
```

This example sets  _R1_ equal to the first footnote reference mark in the active document.




```vb
If ActiveDocument.Footnotes.Count >= 1 Then 
 Set R1 = ActiveDocument.Range.GoTo(What:=wdGoToFootnote, _ 
 Which:=wdGoToFirst) 
 R1.Expand Unit:=wdCharacter 
End If
```


## See also


#### Concepts


[Range Object](range-object-word.md)

