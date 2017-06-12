---
title: Document.GoTo Method (Word)
keywords: vbawd10.chm158007411
f1_keywords:
- vbawd10.chm158007411
ms.prod: word
api_name:
- Word.Document.GoTo
ms.assetid: b03156a8-71a3-af2a-958e-79e1307e1af3
ms.date: 06/08/2017
---


# Document.GoTo Method (Word)

Returns a  **Range** object that represents the start position of the specified item, such as a page, bookmark, or field.


## Syntax

 _expression_ . **GoTo**( **_What_** , **_Which_** , **_Count_** , **_Name_** )

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _What_|Optional| **Variant**|The kind of item to which the range or selection is moved. Can be one of the  **[WdGoToItem](wdgotoitem-enumeration-word.md)** constants.|
| _Which_|Optional| **Variant**|The item to which the range or selection is moved. Can be one of the  **[WdGoToDirection](wdgotodirection-enumeration-word.md)** constants.|
| _Count_|Optional| **Variant**|The number of the item in the document. The default value is 1. Only positive values are valid. To specify an item that precedes the range or selection, use  **wdGoToPrevious** as the Which argument and specify a value for the Count value.|
| _Name_|Optional| **Variant**|If the What argument is  **wdGoToBookmark** , **wdGoToComment** , **wdGoToField** , or **wdGoToObject** , this argument specifies a name. Only positive values are valid. To specify an item that precedes the range or selection, use **wdGoToPrevious** as the Which argument and specify a value for the Count argument.|

## Remarks

When you use the  **GoTo** method with the **wdGoToGrammaticalError** , **wdGoToProofreadingError** , or **wdGoToSpellingError** constant, the **Range** that's returned includes any grammar error text or spelling error text.


## Example

This example sets R1 equal to the first footnote reference mark in the active document.


```vb
If ActiveDocument.Footnotes.Count >= 1 Then 
 Set R1 = ActiveDocument.GoTo(What:=wdGoToFootnote, _ 
 Which:=wdGoToFirst) 
 R1.Expand Unit:=wdCharacter 
End If
```


## See also


#### Concepts


[Document Object](document-object-word.md)

