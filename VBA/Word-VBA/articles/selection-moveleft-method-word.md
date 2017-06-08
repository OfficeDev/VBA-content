---
title: Selection.MoveLeft Method (Word)
keywords: vbawd10.chm158663156
f1_keywords:
- vbawd10.chm158663156
ms.prod: word
api_name:
- Word.Selection.MoveLeft
ms.assetid: 23c22588-e774-f70f-28ea-81b1a54c0dd5
ms.date: 06/08/2017
---


# Selection.MoveLeft Method (Word)

Moves the selection to the left and returns the number of units it has been moved.


## Syntax

 _expression_ . **MoveLeft**( **_Unit_** , **_Count_** , **_Extend_** )

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Unit_|Optional| **WdUnits**|The unit by which the selection is to be moved.The default value is  **wdCharacter** .|
| _Count_|Optional| **Variant**|The number of units the selection is to be moved. The default value is 1.|
| _Extend_|Optional| **Variant**|Can be either  **wdMove** or **wdExtend** . If **wdMove** is used, the selection is collapsed to the endpoint and moved to the left. If **wdExtend** is used, the selection is extended to the left. The default value is **wdMove** .|

## Remarks

When the Unit is  **wdCell** , the Extend argument will only be **wdMove** .


## Example

This example moves the selection one character to the left. If the move is successful, MoveLeft returns 1.


```vb
If Selection.MoveLeft = 1 Then MsgBox "Move was successful"
```

This example enables field shading for the selected field, inserts a DATE field, and then moves the selection left to select the field.




```vb
ActiveDocument.ActiveWindow.View.FieldShading = _ 
 wdFieldShadingWhenSelected 
With Selection 
 .Fields.Add Range:=Selection.Range, Type:=wdFieldDate 
 .MoveLeft Unit:=wdWord, Count:=1 
End With
```

This example moves the selection to the previous table cell.




```vb
If Selection.Information(wdWithInTable) = True Then 
 Selection.MoveLeft Unit:=wdCell, Count:=1, Extend:=wdMove 
End If
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

