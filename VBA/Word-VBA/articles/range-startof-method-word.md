---
title: Range.StartOf Method (Word)
keywords: vbawd10.chm157155435
f1_keywords:
- vbawd10.chm157155435
ms.prod: word
api_name:
- Word.Range.StartOf
ms.assetid: 4d8d5a97-cb5e-cb27-c6dc-35c96c840ae1
ms.date: 06/08/2017
---


# Range.StartOf Method (Word)

Moves or extends the start position of the specified range or selection to the beginning of the nearest specified text unit. This method returns a  **Long** that indicates the number of characters by which the range or selection was moved or extended. The method returns a negative number if the movement is backward through the document.


## Syntax

 _expression_ . **StartOf**( **_Unit_** , **_Extend_** )

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Unit_|Optional| **WdUnits**|The unit by which the start position of the specified range or selection is to be moved. Can be any WdUnits constant except  **wdLine** . The default value is **wdWord** .|
| _Extend_|Optional| **WdMovement**|Specifies whether to move or extend the start of the range. If you use  **wdMove** , both ends of the range or selection are moved to the beginning of the specified unit. If you use **wdExtend** , the beginning of the range or selection is extended to the beginning of the specified unit. The default value is **wdMove** .|

## Remarks

If the beginning of the specified range or selection is already at the beginning of the specified unit, this method doesn't move or extend the range or selection. For example, if the selection is at the beginning of a line, the following example returns 0 (zero) and doesn't change the selection.


```
char = Selection.StartOf(Unit:=wdLine, Extend:=wdMove)
```


## Example

This example selects the text from the insertion point to the beginning of the line. The number of characters selected is stored in  _charmoved_ .


```
Selection.Collapse Direction:=wdCollapseStart charmoved = Selection.StartOf(Unit:=wdLine, Extend:=wdExtend)
```

This example moves the selection to the beginning of the paragraph.




```
Selection.StartOf Unit:=wdParagraph, Extend:=wdMove
```

This example moves  _myRange_ to the beginning of the second sentence in the document ( _myRange_ is collapsed and positioned at the beginning of the second sentence). The example uses the **Select** method to show the location of _myRange_ .




```vb
Set myRange = ActiveDocument.Sentences(2) 
myRange.StartOf Unit:=wdSentence, Extend:=wdMove 
myRange.Select
```


## See also


#### Concepts


[Range Object](range-object-word.md)

