---
title: Selection.EndKey Method (Word)
keywords: vbawd10.chm158663161
f1_keywords:
- vbawd10.chm158663161
ms.prod: word
api_name:
- Word.Selection.EndKey
ms.assetid: 4f27681c-1117-99c2-1aba-bd97082bb8ba
ms.date: 06/08/2017
---


# Selection.EndKey Method (Word)

Moves or extends the selection to the end of the specified unit.


## Syntax

 _expression_ . **EndKey**( **_Unit_** , **_Extend_** )

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Unit_|Optional| **Variant**|The unit by which the selection is to be moved or extended. Can be a  **WdUnits** constant. The default value is **wdLine** .|
| _Extend_|Optional| **Variant**|Specifies the way the selection is moved. Can be any  **WdMovementType** constant. If the value of this argument is **wdMove** , the selection is collapsed to an insertion point and moved to the end of the specified unit. If it is **wdExtend** , the end of the selection is extended to the end of the specified unit. The default value is **wdMove** .|

## Remarks

This method returns an integer that indicates the number of characters the selection or active end was actually moved, or it returns 0 (zero) if the move was unsuccessful. This method corresponds to functionality of the END key.


## Example

This example moves the selection to the end of the current line and assigns the number of characters moved to the pos variable.


```
pos = Selection.EndKey(Unit:=wdLine, Extend:=wdMove)
```

This example moves the selection to the beginning of the current table column and then extends the selection to the end of the column.




```vb
If Selection.Information(wdWithInTable) = True Then 
 Selection.HomeKey Unit:=wdColumn, Extend:=wdMove 
 Selection.EndKey Unit:=wdColumn, Extend:=wdExtend 
End If
```

This example moves the selection to the end of the current story. If the selection is in the main text story, the example moves the selection to the end of the document.




```
Selection.EndKey Unit:=wdStory, Extend:=wdMove
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

