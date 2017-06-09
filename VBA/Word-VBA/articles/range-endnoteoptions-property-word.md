---
title: Range.EndnoteOptions Property (Word)
keywords: vbawd10.chm157155739
f1_keywords:
- vbawd10.chm157155739
ms.prod: word
api_name:
- Word.Range.EndnoteOptions
ms.assetid: 48b2cf9e-edba-e6ed-a3b5-d93e26e17fe5
ms.date: 06/08/2017
---


# Range.EndnoteOptions Property (Word)

Returns an  **EndnoteOptions** object that represents the endnotes in a range.


## Syntax

 _expression_ . **EndnoteOptions**

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


## Example

This example sets the starting number for endnotes in section two of the active document to one if the starting number is not one.


```vb
Sub SetEndnoteOptionsRange() 
 With ActiveDocument.Sections(2).Range.EndnoteOptions 
 If .StartingNumber <> 1 Then 
 .StartingNumber = 1 
 End If 
 End With 
End Sub
```


## See also


#### Concepts


[Range Object](range-object-word.md)

