---
title: Selection.Move Method (Word)
keywords: vbawd10.chm158662765
f1_keywords:
- vbawd10.chm158662765
ms.prod: word
api_name:
- Word.Selection.Move
ms.assetid: 8bd36cf4-ca05-6412-2145-31d07367730e
ms.date: 06/08/2017
---


# Selection.Move Method (Word)

Collapses the specified selection to its start or end position and then moves the collapsed object by the specified number of units. This method returns a  **Long** value that represents the number of units by which the selection was moved, or it returns 0 (zero) if the move was unsuccessful.


## Syntax

 _expression_ . **Move**( **_Unit_** , **_Count_** )

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Unit_|Optional| **[WdUnits](wdunits-enumeration-word.md)**|The unit by which to move the ending character position.|
| _Count_|Optional| **Variant**|The number of units by which the specified range or selection is to be moved. If Count is a positive number, the object is collapsed to its end position and moved backward in the document by the specified number of units. If Count is a negative number, the object is collapsed to its start position and moved forward by the specified number of units. The default value is 1. You can also control the collapse direction by using the  **Collapse** method before using the **Move** method. If the range or selection is in the middle of a unit or isn't collapsed, moving it to the beginning or end of the unit counts as moving it one full unit.|

### Return Value

Long


## Remarks

The start and end positions of a collapsed range or selection are equal.

Applying the  **Move** method to a range doesn't rearrange text in the document. Instead, it redefines the range to refer to a new location in the document.

If you apply the  **Move** method to any range other than a **Range** object variable (for example, `Selection.Paragraphs(3).Range.Move`), the method has no effect.

Moving a  **Selection** object collapses the selection and moves the insertion point either forward or backward in the document.


## Example

This example moves the selection two words to the right and positions the insertion point after the second word's trailing space. If the move is unsuccessful, a message box indicates that the selection is at the end of the document.


```vb
If Selection.StoryType = wdMainTextStory Then 
 wUnits = Selection.Move(Unit:=wdWord, Count:=2) 
 If wUnits < 2 Then _ 
 MsgBox "Selection is at the end of the document" 
End If
```

This example moves the selection forward three cells in the table.




```vb
If Selection.Information(wdWithInTable) = True Then 
 Selection.Move Unit:=wdCell, Count:=3 
End If
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

