---
title: Selection.MoveEndUntil Method (Word)
keywords: vbawd10.chm158662773
f1_keywords:
- vbawd10.chm158662773
ms.prod: word
api_name:
- Word.Selection.MoveEndUntil
ms.assetid: e8f7532a-6a5a-3173-3e5e-db46aec44170
ms.date: 06/08/2017
---


# Selection.MoveEndUntil Method (Word)

Moves the end position of the specified selection until any of the specified characters are found in the document.


## Syntax

 _expression_ . **MoveEndUntil**( **_Cset_** , **_Count_** )

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cset_|Required| **Variant**|One or more characters. This argument is case sensitive.|
| _Count_|Optional| **Variant**|The maximum number of characters by which the specified selection is to be moved. Can be a number or either  **wdForward** or **wdBackward** . If Count is a positive number, the selection is moved forward in the document. If it is a negative number, the selection is moved backward. The default value is **wdForward** .|

### Return Value

Long


## Remarks

This method returns a  **Long** that represents the number of characters by which the end position of the specified selection was moved. If Count is greater than 0 (zero), this method returns the number of characters moved plus 1. If Count is less than 0 (zero), this method returns the number of characters moved minus 1. If no Cset characters are found, the selection isn't changed and the method returns 0 (zero). If the end position is moved backward to a point that precedes the original start position, the start position is set to the new ending position.

 If the movement is forward in the document, the selection is expanded.


## Example

This example extends the selection forward in the document until the letter "a" is found. The example then expands the selection by one character to include the letter "a".


```vb
With Selection 
 .MoveEndUntil Cset:="a", Count:=wdForward 
 .MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend 
End With
```

This example extends the selection forward in the document until a tab is found. If a tab character isn't found in the next 100 characters, the selection isn't moved.




```vb
char = Selection.MoveEndUntil(Cset:=vbTab, Count:=100) 
If char = 0 Then StatusBar = "Selection not moved"
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

