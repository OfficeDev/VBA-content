---
title: Selection.MoveEnd Method (Word)
keywords: vbawd10.chm158662767
f1_keywords:
- vbawd10.chm158662767
ms.prod: word
api_name:
- Word.Selection.MoveEnd
ms.assetid: 11fbcd45-16e6-611b-d296-a88cc7d3ca50
ms.date: 06/08/2017
---


# Selection.MoveEnd Method (Word)

Moves the ending character position of a range or selection.


## Syntax

 _expression_ . **MoveEnd**( **_Unit_** , **_Count_** )

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Unit_|Optional| **WdUnits**|The unit by which to move the ending character position. The default value is  **wdCharacter** .|
| _Count_|Optional| **Variant**|The number of units to move. If this number is positive, the ending character position is moved forward in the document. If this number is negative, the end is moved backward. If the ending position overtakes the starting position, the range collapses and both character positions move together. The default value is 1.|

### Return Value

Integer


## Remarks

This method returns an integer that indicates the number of units the range or selection actually moved, or it returns 0 (zero) if the move was unsuccessful.


## Example

This example moves the end of the selection one character backward (the selection size is reduced by one character). A space is considered a character.


```
Selection.MoveEnd Unit:=wdCharacter, Count:=-1
```

This example moves the end of the selection to the end of the line (the selection is extended to the end of the line).




```
Selection.MoveEnd Unit:=wdLine, Count:=1
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

