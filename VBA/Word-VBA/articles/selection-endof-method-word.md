---
title: Selection.EndOf Method (Word)
keywords: vbawd10.chm158662764
f1_keywords:
- vbawd10.chm158662764
ms.prod: word
api_name:
- Word.Selection.EndOf
ms.assetid: 33aa094b-17f9-3572-f66f-59692c57dc01
ms.date: 06/08/2017
---


# Selection.EndOf Method (Word)

Moves or extends the ending character position of a range or selection to the end of the nearest specified text unit.


## Syntax

 _expression_ . **EndOf**( **_Unit_** , **_Extend_** )

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Unit_|Optional| **Variant**|The unit by which to move the ending character position.  **WdUnits** .|
| _Extend_|Optional| **Variant**|Can be either of the  **WdMovementType** constants. If **wdMove** , both ends of the range or selection object are moved to the end of the specified unit. If **wdExtend** is used, the end of the range or selection is extended to the end of the specified unit. The default value is **wdMove** .|

## Remarks

This method returns a value that indicates the number of character positions the range or selection was moved or extended (movement is forward in the document).

If the both the starting and ending positions for the range or selection are already at the end of the specified unit, this method doesn't move or extend the range or selection. For example, if the selection is at the end of a word and the trailing space, the following instruction doesn't change the selection ( _char_ equals 0 (zero)).




```
char = Selection.EndOf(Unit:=wdWord, Extend:=wdMove)
```


## Example

This example extends the selection to the end of the paragraph.


```vb
charmoved = Selection.EndOf(Unit:=wdParagraph, Extend:=wdExtend) 
If charmoved = 0 Then MsgBox "Selection unchanged"
```

This example moves myRange to the end of the first word in the selection (after the trailing space).




```vb
Set myRange = Selection.Characters(1) 
myRange.EndOf Unit:=wdWord, Extend:=wdMove
```

This example adds a table, selects the first cell in row two, and then extends the selection to the end of the column.




```vb
Set myRange = ActiveDocument.Range(0, 0) 
Set myTable = ActiveDocument.Tables.Add(Range:=myRange, _ 
 NumRows:=5, NumColumns:=3) 
myTable.Cell(2, 1).Select 
Selection.EndOf Unit:=wdColumn, Extend:=wdExtend
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

