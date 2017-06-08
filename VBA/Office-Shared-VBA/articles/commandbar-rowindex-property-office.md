---
title: CommandBar.RowIndex Property (Office)
keywords: vbaof11.chm3014
f1_keywords:
- vbaof11.chm3014
ms.prod: office
api_name:
- Office.CommandBar.RowIndex
ms.assetid: 6dd5576c-0a46-9a72-9c4e-fcf685097b77
ms.date: 06/08/2017
---


# CommandBar.RowIndex Property (Office)

Gets or sets the docking order of a command bar in relation to other command bars in the same docking area. Can be an integer greater than zero, or either of the following  **MsoBarRow** constants: **msoBarRowFirst** or **msoBarRowLast**. Read/write.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **RowIndex**

 _expression_ A variable that represents a **CommandBar** object.


## Remarks

Several command bars can share the same row index, and command bars with lower numbers are docked first. If two or more command bars share the same row index, the command bar most recently assigned will be displayed first in its group.


## Example

This example adjusts the position of the command bar named "Custom" by moving it to the left 110 pixels more than the default, and it makes this command bar the first to be docked by changing its row index to  **msoBarRowFirst**.


```
Set myBar = CommandBars("Custom") 
With myBar 
    .RowIndex = msoBarRowFirst 
    .Left = 140 
End With
```


## See also


#### Concepts


[CommandBar Object](commandbar-object-office.md)
#### Other resources


[CommandBar Object Members](commandbar-members-office.md)

