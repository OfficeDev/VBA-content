---
title: Selection.MoveEndWhile Method (Word)
keywords: vbawd10.chm158662770
f1_keywords:
- vbawd10.chm158662770
ms.prod: word
api_name:
- Word.Selection.MoveEndWhile
ms.assetid: c1cd97cc-9836-a61f-3b7a-4e178bc3d1e0
ms.date: 06/08/2017
---


# Selection.MoveEndWhile Method (Word)

Moves the ending character position of a selection while any of the specified characters are found in the document.


## Syntax

 _expression_ . **MoveEndWhile**( **_Cset_** , **_Count_** )

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cset_|Required| **Variant**|One or more characters. This argument is case sensitive.|
| _Count_|Optional| **Variant**|The maximum number of characters by which the selection is to be moved. Can be a number or either  **wdForward** or **wdBackward** . If Count is a positive number, the selection is moved forward in the document. If it is a negative number, the selection is moved backward. The default value is **wdForward** .|

## Remarks

While any character in Cset is found, the end position of the specified selection is moved. This method returns the number of characters that the end position of the selection moved as a  **Long** value. If no Cset characters are found, the selection isn't changed and the method returns 0 (zero). If the end position is moved backward to a point that precedes the original start position, the start position is set to the new end position.


## Example

This example moves the end position of the selection forward while the space character is found.


```
Selection.MoveEndWhile Cset:=" ", Count:=wdForward
```

This example moves the end position of the selection forward while  **Count** is less than or equal to 10 and any letter from "a" through "h" is found.




```
Selection.MoveEndWhile Cset:="abcdefgh", Count:=10
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

