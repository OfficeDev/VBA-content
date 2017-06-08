---
title: Range.SetRange Method (Word)
keywords: vbawd10.chm157155428
f1_keywords:
- vbawd10.chm157155428
ms.prod: word
api_name:
- Word.Range.SetRange
ms.assetid: 91097079-406c-98f4-d37c-cca8dab7aef0
ms.date: 06/08/2017
---


# Range.SetRange Method (Word)

Sets the starting and ending character positions for an existing range.


## Syntax

 _expression_ . **SetRange**( **_Start_** , **_End_** )

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Start_|Required| **Long**|The starting character position of the range.|
| _End_|Required| **Long**|The ending character position of the range.|

## Remarks

Character position values start at the beginning of the story, with the first value being 0 (zero). All characters are counted, including nonprinting characters. Hidden characters are counted even if they're not displayed.

The  **SetRange** method redefines the starting and ending positions of an existing **Range** object. This method differs from the **Range** method, which is used to create a range, given a starting and ending position.


## Example

This example uses  **SetRange** to redefine _myRange_ to refer to the first three paragraphs in the active document.


```vb
Set myRange = ActiveDocument.Paragraphs(1).Range 
myRange.SetRange Start:=myRange.Start, _ 
 End:=ActiveDocument.Paragraphs(3).Range.End
```

This example uses  **SetRange** to redefine _myRange_ to refer to the area starting at the beginning of the document and ending at the end of the current selection.




```vb
Set myRange = ActiveDocument.Range(Start:=0, End:=0) 
myRange.InsertAfter "Hello " 
myRange.SetRange Start:=myRange.Start, End:=Selection.End
```


## See also


#### Concepts


[Range Object](range-object-word.md)

