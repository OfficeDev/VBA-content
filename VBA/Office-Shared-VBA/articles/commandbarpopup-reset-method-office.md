---
title: CommandBarPopup.Reset Method (Office)
ms.prod: office
api_name:
- Office.CommandBarPopup.Reset
ms.assetid: 8e31b4e2-66d1-b902-f837-dc4833b1607f
ms.date: 06/08/2017
---


# CommandBarPopup.Reset Method (Office)

Resets a built-in  **CommandBarPopup** control to its original function and face.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **Reset**

 _expression_ A variable that represents a **CommandBarPopup** object.


## Remarks

Resetting a built-in control restores the actions originally intended for the control and resets each of the control's properties back to its original state. 


## Example

The following example searches all command bars for a CommandBarPopup object whose tag is "Graphics" and then resets it to its default state.


```
Set myControl = Application.CommandBars.FindControl _ 
(Type:=msoControlPopup, Tag:="Graphics")  
myControl.Reset 

```


## See also


#### Concepts


[CommandBarPopup Object](commandbarpopup-object-office.md)
#### Other resources


[CommandBarPopup Object Members](commandbarpopup-members-office.md)

