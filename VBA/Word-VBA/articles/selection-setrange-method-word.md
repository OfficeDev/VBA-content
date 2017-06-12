---
title: Selection.SetRange Method (Word)
keywords: vbawd10.chm158662756
f1_keywords:
- vbawd10.chm158662756
ms.prod: word
api_name:
- Word.Selection.SetRange
ms.assetid: 232a681e-4205-05ae-f442-9dc1a2df96f1
ms.date: 06/08/2017
---


# Selection.SetRange Method (Word)

Sets the starting and ending character positions for the selection.


## Syntax

 _expression_ . **SetRange**( **_Start_** , **_End_** )

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Start_|Required| **Long**|The starting character position of the selection.|
| _End_|Required| **Long**|The ending character position of the selection.|

## Remarks

Character position values start at the beginning of the story, with the first value being 0 (zero). All characters are counted, including nonprinting characters. Hidden characters are counted even if they're not displayed.

The  **SetRange** method redefines the starting and ending positions of an existing **Selection** object. This method differs from the **Range** method, which is used to create a **Range** object, given a starting and ending position.


## Example

This example selects the first 10 characters in the document.


```
Selection.SetRange Start:=0, End:=10
```

This example extends the selection to the end of the document.




```
Selection.SetRange Start:=Selection.Start, _ 
 End:=ActiveDocument.Content.End
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

