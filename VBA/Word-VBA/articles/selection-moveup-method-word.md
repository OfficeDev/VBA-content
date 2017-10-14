---
title: Selection.MoveUp Method (Word)
keywords: vbawd10.chm158663158
f1_keywords:
- vbawd10.chm158663158
ms.prod: word
api_name:
- Word.Selection.MoveUp
ms.assetid: 46993371-c916-06b5-a644-960f8a283536
ms.date: 06/08/2017
---


# Selection.MoveUp Method (Word)

Moves the selection up and returns the number of units that it has been moved.


## Syntax

 _expression_ . **MoveUp**( **_Unit_** , **_Count_** , **_Extend_** )

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Unit_|Optional| **Variant**|The unit by which to move the selection. Can be one of the following  **WdUnits** constants: **wdLine** , **wdParagraph** , **wdWindow** or **wdScreen** . The default value is **wdLine** . You can use the **wdWindow** constant for the Unit argument to move to the top or bottom of the active window. Regardless of the value of Count (greater than 1 or less than ? 1), the **wdWindow** constant moves only one unit. Use the **wdScreen** constant to move more than one screen.|
| _Count_|Optional| **Variant**|The number of units the selection is to be moved. The default value is 1.|
| _Extend_|Optional| **Variant**|Specifies whether the selection is moved or extended. Can be either  **wdMove** or **wdExtend** . If **wdMove** is used, the selection is collapsed to the endpoint and moved up. If **wdExtend** is used, the selection is extended up. The default value is **wdMove** .|

### Return Value

Long


## Example

This example moves the selection to the beginning of the previous paragraph.


```
Selection.MoveRight 
Selection.MoveUp Unit:=wdParagraph, Count:=2, Extend:=wdMove
```

This example displays the current line number, moves the selection up three lines, and displays the current line number again.




```vb
MsgBox "Line " &; Selection.Information(wdFirstCharacterLineNumber) 
Selection.MoveUp Unit:=wdLine, Count:=3, Extend:=wdMove 
MsgBox "Line " &; Selection.Information(wdFirstCharacterLineNumber)
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

