---
title: Selection.HomeKey Method (Word)
keywords: vbawd10.chm158663160
f1_keywords:
- vbawd10.chm158663160
ms.prod: word
api_name:
- Word.Selection.HomeKey
ms.assetid: 24264193-d610-acbc-b393-de41fd55e976
ms.date: 06/08/2017
---


# Selection.HomeKey Method (Word)

Moves or extends the selection to the beginning of the specified unit. This method returns an integer that indicates the number of characters the selection was actually moved, or it returns 0 (zero) if the move was unsuccessful.This method corresponds to functionality of the HOME key.


## Syntax

 _expression_ . **HomeKey**( **_Unit_** , **_Extend_** )

 _expression_ A variable that represents a **[Selection](selection-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Unit_|Optional| **Variant**|The unit by which the selection is to be moved or extended. The default value is  **wdLine** .|
| _Extend_|Optional| **Variant**|Specifies the way the selection is moved. Can be one of the  **WdMovementType** constants. If the value of this argument is **wdMove** , the selection is collapsed to an insertion point and moved to the beginning of the specified unit. If it is **wdExtend** , the beginning of the selection is extended to the beginning of the specified unit. The default value is **wdMove** .|

## Example

This example moves the selection to the beginning of the current story. If the selection is in the main text story, the selection is moved to the beginning of the document.


```
Selection.HomeKey Unit:=wdStory, Extend:=wdMove
```

This example moves the selection to the beginning of the current line and assigns the number of characters moved to the pos variable.




```vb
pos = Selection.HomeKey(Unit:=wdLine, Extend:=wdMove) 
If pos = 0 Then StatusBar = "Selection was not moved"
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

