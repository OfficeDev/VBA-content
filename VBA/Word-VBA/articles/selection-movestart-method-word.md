---
title: Selection.MoveStart Method (Word)
keywords: vbawd10.chm158662766
f1_keywords:
- vbawd10.chm158662766
ms.prod: word
api_name:
- Word.Selection.MoveStart
ms.assetid: c58f4dd5-791b-ac0f-8445-29e0ade48d7f
ms.date: 06/08/2017
---


# Selection.MoveStart Method (Word)

Moves the start position of the specified selection.


## Syntax

 _expression_ . **MoveStart**( **_Unit_** , **_Count_** )

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Unit_|Optional| **WdUnits**|The unit by which start position of the specified selection is to be moved. Can be one of the  **WdUnits** constants. The default value is **wdCharacter** .|
| _Count_|Optional| **Variant**|The maximum number of units by which the specified selection is to be moved. If Count is a positive number, the start position of the selection is moved forward in the document. If it is a negative number, the start position is moved backward. If the start position is moved forward to a position beyond the end position, the selection is collapsed and both the start and end positions are moved together. The default value is 1.|

### Return Value

Integer


## Remarks

This method returns an integer that indicates the number of units by which the start position or the selection actually moved, or it returns 0 (zero) if the move was unsuccessful.


## Example

This example moves the start position of the selection one character forward (the selection size is reduced by one character). Note that a space is considered a character.


```
Selection.MoveStart Unit:=wdCharacter, Count:=1
```

This example moves the start position of the selection to the beginning of the line (the selection is extended to the start of the line).




```
Selection.MoveStart Unit:=wdLine, Count:=-1
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

