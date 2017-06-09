---
title: ListFormat.ApplyOutlineNumberDefault Method (Word)
keywords: vbawd10.chm163578070
f1_keywords:
- vbawd10.chm163578070
ms.prod: word
api_name:
- Word.ListFormat.ApplyOutlineNumberDefault
ms.assetid: 8d3d26ad-e01c-8ad4-d4f4-86e71628e2c3
ms.date: 06/08/2017
---


# ListFormat.ApplyOutlineNumberDefault Method (Word)

Adds the default outline-numbering scheme to the paragraphs in the range for the specified  **ListFormat** object.


## Syntax

 _expression_ . **ApplyOutlineNumberDefault**( **_DefaultListBehavior_** )

 _expression_ Required. A variable that represents a **[ListFormat](listformat-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DefaultListBehavior_|Optional| **Variant**|Sets a value that specifies whether Microsoft Word uses new Web-oriented formatting for better list display. Can be either of the following constants:  **wdWord8ListBehavior** (use formatting compatible with Microsoft Word 97) or **wdWord9ListBehavior** (use Web-oriented formatting). For compatibility reasons, the default constant is **wdWord8ListBehavior** , but in new procedures you should use **wdWord9ListBehavior** to take advantage of improved Web-oriented formatting with respect to indenting and multilevel lists.|

## Remarks

If the paragraphs are already formatted as an outline-numbered list, this method removes the numbers and formatting. This method doesn't remove built-in heading styles that have been applied to paragraphs.


## Example

This example adds outline numbering to the paragraphs in the selection. If the selection is already an outline-numbered list, the example removes the numbers and formatting.


```
Selection.Range.ListFormat.ApplyOutlineNumberDefault
```

This example sets the variable myRange to include paragraphs three through six of the active document, and then it checks to see whether the range contains list formatting. If there is no list formatting, the default outline-numbered list format is applied.




```vb
Set myDoc = ActiveDocument 
Set myRange = myDoc.Range( _ 
 Start:= myDoc.Paragraphs(3).Range.Start, _ 
 End:=myDoc.Paragraphs(6).Range.End) 
If myRange.ListFormat.ListType = wdListNoNumbering Then 
 myRange.ListFormat.ApplyOutlineNumberDefault 
End If
```


## See also


#### Concepts


[ListFormat Object](listformat-object-word.md)

