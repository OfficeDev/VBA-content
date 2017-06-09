---
title: ListFormat.ListValue Property (Word)
keywords: vbawd10.chm163577927
f1_keywords:
- vbawd10.chm163577927
ms.prod: word
api_name:
- Word.ListFormat.ListValue
ms.assetid: 58c07741-b59f-60c0-bff1-0a63eb61847c
ms.date: 06/08/2017
---


# ListFormat.ListValue Property (Word)

Returns the numeric value of the first paragraph in the range for the specified  **ListFormat** object. Read-only **Long** .


## Syntax

 _expression_ . **ListValue**

 _expression_ An expression that returns a **ListFormat** object.


## Remarks

Use the  **ListString** property to return a string that represents the list value.

If the  **ListFormat** object applies to a bulleted list, the **ListValue** property returns 1.

If the  **ListFormat** object applies to an outline-numbered list, the **ListValue** property returns the numeric value of the first paragraph as it occurs in the sequence of paragraphs at the same level. For example, if the first paragraph for a specified ListFormat object were numbered "A.2," the **ListValue** property would return 2.

This property will not return the value for a LISTNUM field.


## Example

This example displays both the numeric value of the first paragraph in the selection and the string representation of that value.


```
v = Selection.Range.ListFormat.ListValue 
lstring = Selection.Range.ListFormat.ListString 
MsgBox "List value " &; v _ 
 &; " is represented by the string " &; lstring
```


## See also


#### Concepts


[ListFormat Object](listformat-object-word.md)

