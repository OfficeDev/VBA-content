---
title: Selection.InRange Method (Word)
keywords: vbawd10.chm158662782
f1_keywords:
- vbawd10.chm158662782
ms.prod: word
api_name:
- Word.Selection.InRange
ms.assetid: 3759ad96-44b5-d63c-f4d5-844f937f4216
ms.date: 06/08/2017
---


# Selection.InRange Method (Word)

 **True** if the selection to which the method is applied is contained within the range specified by the Range argument.


## Syntax

 _expression_ . **InRange**( **_Range_** )

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Range_|Required|Optional|The [Range](range-object-word.md) to which you want to compare the selection.|

### Return Value

Boolean


## Remarks

This method determines whether the range or selection returned by expression is contained in the specified Range by comparing the starting and ending character positions and the story type.


## Example

This example determines whether the selection is contained in the first paragraph in the active document.


```
status = Selection.InRange(ActiveDocument.Paragraphs(1).Range)
```

This example displays a message if the selection is in the footnote story.




```vb
If Selection.InRange(ActiveDocument _ 
 .StoryRanges(wdFootnotesStory)) Then 
 MsgBox "Selection in footnotes" 
End If
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

