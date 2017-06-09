---
title: Range.EndOf Method (Word)
keywords: vbawd10.chm157155436
f1_keywords:
- vbawd10.chm157155436
ms.prod: word
api_name:
- Word.Range.EndOf
ms.assetid: b9bda3b3-fee5-6359-f0ab-10fbe6b635b1
ms.date: 06/08/2017
---


# Range.EndOf Method (Word)

Moves or extends the ending character position of a range to the end of the nearest specified text unit.


## Syntax

 _expression_ . **EndOf**( **_Unit_** , **_Extend_** )

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Unit_|Optional| **Variant**|The unit by which to move the ending character position. Can be any  **WdUnits** , except **wdLine** . The default value is **wdWord** .|
| _Extend_|Required| ** WdMovementType**|Specifies whether to move or extend the end of the range. If the value is  **wdMove** , both ends of the range or selection object are moved to the end of the specified unit. If **wdExtend** is used, the end of the range or selection is extended to the end of the specified unit. The default value is **wdMove** .|

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


[Range Object](range-object-word.md)

