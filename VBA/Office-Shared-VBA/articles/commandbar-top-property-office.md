---
title: CommandBar.Top Property (Office)
keywords: vbaof11.chm3018
f1_keywords:
- vbaof11.chm3018
ms.prod: office
api_name:
- Office.CommandBar.Top
ms.assetid: 1bac668a-0caa-d185-cc07-ba55809c79fe
ms.date: 06/08/2017
---


# CommandBar.Top Property (Office)

Sets or gets the distance from the top edge of the specified command bar, to the top edge of the screen. For docked command bars, this property returns or sets the distance from the command bar to the top of the docking area. Read/write.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **Top**

 _expression_ Required. A variable that represents a **[CommandBar](commandbar-object-office.md)** object.


## Example

This example positions the upper-left corner of the floating command bar named "Custom" 140 pixels from the left edge of the screen and 100 pixels from the top of the screen.


```
Set myBar = CommandBars("Custom") 
myBar.Position = msoBarFloating 
With myBar 
    .Left = 140 
    .Top = 100 
End With
```


## See also


#### Concepts


[CommandBar Object](commandbar-object-office.md)
#### Other resources


[CommandBar Object Members](commandbar-members-office.md)

