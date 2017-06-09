---
title: CommandBarPopup.SetFocus Method (Office)
ms.prod: office
api_name:
- Office.CommandBarPopup.SetFocus
ms.assetid: ce132a0d-aa1f-c8b1-2697-1cfe78b99123
ms.date: 06/08/2017
---


# CommandBarPopup.SetFocus Method (Office)

Moves the keyboard focus to the specified  **CommandBarPopup** control. If the popup is disabled or isn't visible, this method will fail.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **SetFocus**

 _expression_ A variable that represents a **CommandBarPopup** object.


## Example

The following example sets a reference to an existing command bar popup and then resets it to its default state.


```
Dim cbPopup As CommandBarPopup 
Set cbPopup = Application.CommandBars.FindControl _ 
(Type:=msoControlPopup, Tag:="Graphics") 
cbPopup.Reset 

```


## See also


#### Concepts


[CommandBarPopup Object](commandbarpopup-object-office.md)
#### Other resources


[CommandBarPopup Object Members](commandbarpopup-members-office.md)

