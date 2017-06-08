---
title: Selection.MoveStartUntil Method (Word)
keywords: vbawd10.chm158662772
f1_keywords:
- vbawd10.chm158662772
ms.prod: word
api_name:
- Word.Selection.MoveStartUntil
ms.assetid: a461cf49-1ed9-425b-5417-0a882c17d792
ms.date: 06/08/2017
---


# Selection.MoveStartUntil Method (Word)

Moves the start position of the specified selection until one of the specified characters is found in the document. If the movement is backward through the document, the selection is expanded.


## Syntax

 _expression_ . **MoveStartUntil**( **_Cset_** , **_Count_** )

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cset_|Required| **Variant**|One or more characters. This argument is case sensitive.|
| _Count_|Optional| **Variant**|The maximum number of characters by which the specified selection is to be moved. Can be a number or either the  **wdForward** or **wdBackward** constant. If Count is a positive number, the selection is moved forward in the document. If it is a negative number, the selection is moved backward. The default value is **wdForward** .|

## Remarks

This method returns the number of characters by which the start position of the specified selection moved, as a  **Long** value. If Count is greater than 0 (zero), this method returns the number of characters moved plus 1. If Count is less than 0 (zero), this method returns the number of characters moved minus 1. If no Cset characters are found, the specified selection isn't changed and the method returns 0 (zero). If the start position is moved forward to a point beyond the end position, the selection is collapsed and both the start and end positions are moved together.


## Example

This example extends the selection backward until a capital "I" is found.


```
Selection.MoveStartUntil Cset:="I", Count:=wdBackward
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

