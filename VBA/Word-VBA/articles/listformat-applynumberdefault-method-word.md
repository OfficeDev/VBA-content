---
title: ListFormat.ApplyNumberDefault Method (Word)
keywords: vbawd10.chm163578069
f1_keywords:
- vbawd10.chm163578069
ms.prod: word
api_name:
- Word.ListFormat.ApplyNumberDefault
ms.assetid: de7e219c-fb92-b0cf-dbc0-33f98eee0f5a
ms.date: 06/08/2017
---


# ListFormat.ApplyNumberDefault Method (Word)

Adds the default numbering scheme to the paragraphs in the range for the specified  **ListFormat** object.


## Syntax

 _expression_ . **ApplyNumberDefault**( **_DefaultListBehavior_** )

 _expression_ Required. A variable that represents a **[ListFormat](listformat-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DefaultListBehavior_|Optional| **Variant**|Sets a value that specifies whether Microsoft Word uses new Web-oriented formatting for better list display. Can be either of the following constants:  **wdWord8ListBehavior** (use formatting compatible with Microsoft Word 97) or **wdWord9ListBehavior** (use Web-oriented formatting). For compatibility reasons, the default constant is **wdWord8ListBehavior** , but in new procedures you should use **wdWord9ListBehavior** to take advantage of improved Web-oriented formatting with respect to indenting and multilevel lists.|

## Remarks

If the paragraphs are already formatted as a numbered list, this method removes the numbers and formatting.


## Example

This example numbers the paragraphs in the selection. If the selection is already a numbered list, the example removes the numbers and formatting.


```
Selection.Range.ListFormat.ApplyNumberDefault
```

This example sets the variable myRange to include paragraphs three through six of the active document, and then it checks to see whether the range contains list formatting. If there is no list formatting, default numbers are applied to the range.




```vb
Set myDoc = ActiveDocument 
Set myRange = myDoc.Range( _ 
 Start:= myDoc.Paragraphs(3).Range.Start, _ 
 End:=myDoc.Paragraphs(6).Range.End) 
If myRange.ListFormat.ListType = wdListNoNumbering Then 
 myRange.ListFormat.ApplyNumberDefault 
End If
```


## See also


#### Concepts


[ListFormat Object](listformat-object-word.md)

