---
title: Selection.MoveStartWhile Method (Word)
keywords: vbawd10.chm158662769
f1_keywords:
- vbawd10.chm158662769
ms.prod: word
api_name:
- Word.Selection.MoveStartWhile
ms.assetid: b6e33ffc-a07f-2ef9-0e35-55aaf256f098
ms.date: 06/08/2017
---


# Selection.MoveStartWhile Method (Word)

Moves the start position of the specified selection while any of the specified characters are found in the document.


## Syntax

 _expression_ . **MoveStartWhile**( **_Cset_** , **_Count_** )

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cset_|Required| **Variant**|One or more characters. This argument is case sensitive.|
| _Count_|Optional| **Variant**|The maximum number of characters by which the specified selection is to be moved. Can be a number or either the  **wdForward** or **wdBackward** constant. If Count is a positive number, the selection is moved forward in the document. If it is a negative number, the selection is moved backward. The default value is **wdForward** .|

## Remarks

While any character in Cset is found, the start position of the selection is moved. This method returns the number of characters that the start position of the selection moved as a  **Long** value. If not Cset characters are found, the selection isn't changed and the method returns 0 (zero). If the start position is moved forward to a position beyond the original end position, the end position is set to the new start position.


## Example

This example moves the start position of the selection backward through the document while the space character is found.


```
Selection.MoveStartWhile Cset:=" ", Count:=wdBackward
```

This example moves the start position of the selection backward through the document while Count is less than or equal to 10 and any letter from "a" through "h" is found.




```
Selection.MoveStartWhile Cset:="abcdefgh", Count:=-10
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

