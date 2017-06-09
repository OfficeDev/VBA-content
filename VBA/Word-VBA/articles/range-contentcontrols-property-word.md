---
title: Range.ContentControls Property (Word)
keywords: vbawd10.chm157155752
f1_keywords:
- vbawd10.chm157155752
ms.prod: word
api_name:
- Word.Range.ContentControls
ms.assetid: e8c715af-067f-871e-7dec-28aa4302d9f9
ms.date: 06/08/2017
---


# Range.ContentControls Property (Word)

Returns a  **[ContentControls](contentcontrols-object-word.md)** collection that represents the content controls contained within a range. Read-only.


## Syntax

 _expression_ . **ContentControls**

 _expression_ An expression that returns a **[Range](range-object-word.md)** object.


## Example

The following example inserts a drop-down list content control into the active document at the specified position.


```vb
Dim objCC As ContentControl 
Dim objRange as Range 
 
Set objRange = ActiveDocument.Range(200, 200) 
Set objCC = objRange.ContentControls.Add(wdContentControlDropdownList) 
 
'List entries 
objCC.DropdownListEntries.Add "Cat" 
objCC.DropdownListEntries.Add "Dog" 
objCC.DropdownListEntries.Add "Horse" 
objCC.DropdownListEntries.Add "Monkey" 
objCC.DropdownListEntries.Add "Snake" 
objCC.DropdownListEntries.Add "Other"
```


## See also


#### Concepts


[Range Object](range-object-word.md)

