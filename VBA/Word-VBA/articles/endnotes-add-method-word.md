---
title: Endnotes.Add Method (Word)
keywords: vbawd10.chm155254788
f1_keywords:
- vbawd10.chm155254788
ms.prod: word
api_name:
- Word.Endnotes.Add
ms.assetid: 6931462d-ee52-862b-3c63-127ebc828c5e
ms.date: 06/08/2017
---


# Endnotes.Add Method (Word)

Returns an  **Endnote** object that represents an endnote added to a range.


## Syntax

 _expression_ . **Add**( **_Range_** , **_Reference_** , **_Text_** )

 _expression_ Required. A variable that represents an **[Endnotes](endnotes-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range object**|The range marked for the endnote or footnote. This can be a collapsed range.|
| _Reference_|Optional| **Variant**|The text for the custom reference mark. If this argument is omitted, Microsoft Word inserts an automatically-numbered reference mark.|
| _Text_|Optional| **Variant**|The text of the endnote or footnote.|

## Remarks

To specify a symbol for the Reference argument, use the syntax  `{FontName CharNum}`. FontName is the name of the font that contains the symbol. Names of decorative fonts appear in the  **Font** box in the **Symbol** dialog box ( **Insert** menu). CharNum is the sum of 31 and the number corresponding to the position of the symbol you want to insert, counting from left to right in the table of symbols. For example, to specify an omega symbol at position 56 in the table of symbols in the Symbol font, the argument would be "{Symbol 87}".


## Example

This example adds an endnote to the third paragraph in the active document


```vb
Set myRange = ActiveDocument.Paragraphs(3).Range 
ActiveDocument.Endnotes.Add Range:=myRange, _ 
 Text:= "Ibid., 314."
```


## See also


#### Concepts


[Endnotes Collection Object](endnotes-object-word.md)

