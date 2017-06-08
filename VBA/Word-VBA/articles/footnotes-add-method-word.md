---
title: Footnotes.Add Method (Word)
keywords: vbawd10.chm155320324
f1_keywords:
- vbawd10.chm155320324
ms.prod: word
api_name:
- Word.Footnotes.Add
ms.assetid: 952a90b0-f550-820b-15e7-82bad3cc201f
ms.date: 06/08/2017
---


# Footnotes.Add Method (Word)

Returns a  **Footnote** object that represents a footnote added to a range.


## Syntax

 _expression_ . **Add**( **_Range_** , **_Reference_** , **_Text_** )

 _expression_ Required. A variable that represents a **[Footnotes](footnotes-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range object**|The range marked for the endnote or footnote. This can be a collapsed range.|
| _Reference_|Optional| **Variant**|The text for the custom reference mark. If this argument is omitted, Microsoft Word inserts an automatically-numbered reference mark.|
| _Text_|Optional| **Variant**|The text of the endnote or footnote.|

### Return Value

Footnote


## Remarks

To specify a symbol for the Reference argument, use the syntax  `{FontName CharNum}`. FontName is the name of the font that contains the symbol. Names of decorative fonts appear in the  **Font** box in the **Symbol** dialog box ( **Insert** menu). CharNum is the sum of 31 and the number corresponding to the position of the symbol you want to insert, counting from left to right in the table of symbols. For example, to specify an omega symbol at position 56 in the table of symbols in the Symbol font, the argument would be "{Symbol 87}".


## Example

The following code example adds an automatically-numbered footnote at the end of the selection.


```vb
ActiveDocument.Footnotes.Add Range:= Selection.Range , _ 
 Text:= "The Willow Tree, (Lone Creek Press, 1996)."
```

The following code example adds a footnote that uses a custom symbol for the reference mark.




```vb
ActiveDocument.Footnotes.Add Range:= Selection.Range , _ 
 Text:= "More information in the full report.", _ 
 Reference:= "{Symbol -3998}"
```


## See also


#### Concepts


[Footnotes Collection Object](footnotes-object-word.md)

