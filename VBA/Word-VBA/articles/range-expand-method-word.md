---
title: Range.Expand Method (Word)
keywords: vbawd10.chm157155457
f1_keywords:
- vbawd10.chm157155457
ms.prod: word
api_name:
- Word.Range.Expand
ms.assetid: cf4a5705-ebda-fedb-4929-3e115d42a432
ms.date: 06/08/2017
---


# Range.Expand Method (Word)

Expands the specified range or selection. Returns the number of characters added to the range or selection.  **Long** .


## Syntax

 _expression_ . **Expand**( **_Unit_** )

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Unit_|Optional| **Variant**|The unit by which to expand the range. Can be one of the following  **WdUnits** constants: **wdCharacter** , **wdWord** , **wdSentence** , **wdParagraph** , **wdSection** , **wdStory** , **wdCell** **wdColumn** , **wdRow** , or **wdTable** .|

## Example

This example creates a range that refers to the first word in the active document, and then it expands the range to reference the first paragraph in the document.


```vb
Set myRange = ActiveDocument.Words(1) 
myRange.Expand Unit:=wdParagraph
```

This example capitalizes the first character in the selection and then expands the selection to include the entire sentence.




```vb
With Selection 
 .Characters(1).Case = wdTitleSentence 
 .Expand Unit:=wdSentence 
End With
```


## See also


#### Concepts


[Range Object](range-object-word.md)

