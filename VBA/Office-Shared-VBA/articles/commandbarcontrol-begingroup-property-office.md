---
title: CommandBarControl.BeginGroup Property (Office)
ms.prod: office
api_name:
- Office.CommandBarControl.BeginGroup
ms.assetid: 529b8c23-ec1f-b37b-a40c-9ae6016f4dc0
ms.date: 06/08/2017
---


# CommandBarControl.BeginGroup Property (Office)

Gets  **True** if the specified command bar control appears at the beginning of a group of controls on the command bar. Read/write.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **BeginGroup**

 _expression_ A variable that represents a **CommandBarControl** object.


### Return Value

Boolean


## Example

This example begins a new group with the last control on the active menu bar.


```
Set myMenuBar = CommandBars.ActiveMenuBar 
Set lastMenu = myMenuBar _ 
    .Controls(myMenuBar.Controls.Count) 
lastMenu.BeginGroup = True
```


## See also


#### Concepts


[CommandBarControl Object](commandbarcontrol-object-office.md)
#### Other resources


[CommandBarControl Object Members](commandbarcontrol-members-office.md)

