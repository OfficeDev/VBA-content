---
title: Range.Move Method (Word)
keywords: vbawd10.chm157155437
f1_keywords:
- vbawd10.chm157155437
ms.prod: word
api_name:
- Word.Range.Move
ms.assetid: 40c73c63-12da-4e8c-05c3-121f4df57f3f
ms.date: 06/08/2017
---


# Range.Move Method (Word)

Collapses the specified range to its start or end position and then moves the collapsed object by the specified number of units.


## Syntax

 _expression_ . **Move**( **_Unit_** , **_Count_** )

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Unit_|Optional| **WdUnits**|The unit by which to move the range.|
| _Count_|Optional| **Variant**|The number of units by which the specified range is to be moved. If Count is a positive number, the object is collapsed to its end position and moved backward in the document by the specified number of units. If Count is a negative number, the object is collapsed to its start position and moved forward by the specified number of units. The default value is 1. You can also control the collapse direction by using the  **Collapse** method before using the **Move** method. If the range is in the middle of a unit or isn't collapsed, moving it to the beginning or end of the unit counts as moving it one full unit.|

### Return Value

Long


## Remarks

This method returns a  **Long** value that indicates the number of units by which the object was actually moved, or it returns 0 (zero) if the move was unsuccessful.

The start and end positions of a collapsed range are equal.

Applying the  **Move** method to a range doesn't rearrange text in the document. Instead, it redefines the range to refer to a new location in the document.


## Example

This example sets  _Range1_ to the first paragraph in the active document and then moves the range forward three paragraphs. After this macro is run, the insertion point is positioned at the beginning of the fourth paragraph.


```vb
Set Range1 = ActiveDocument.Paragraphs(1).Range 
With Range1 
 .Collapse Direction:=wdCollapseStart 
 .Move Unit:=wdParagraph, Count:=3 
 .Select 
End With
```


## See also


#### Concepts


[Range Object](range-object-word.md)

