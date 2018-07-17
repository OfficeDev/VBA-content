---
title: CommandBar.Reset Method (Office)
keywords: vbaof11.chm3016
f1_keywords:
- vbaof11.chm3016
ms.prod: office
api_name:
- Office.CommandBar.Reset
ms.assetid: 96dfb3cc-a53c-ea7f-eb98-96a983faa681
ms.date: 06/08/2017
---


# CommandBar.Reset Method (Office)

Resets a built-in command bar to its default configuration.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **Reset**

 _expression_ A variable that represents a **CommandBar** object.


## Remarks

Resetting a built-in control restores the actions originally intended for the control and resets each of the control's properties back to its original state. Resetting a built-in command bar removes custom controls and restores built-in controls.


## Example

This example uses the value of User to adjust the command bars according to the user level. If User is "Level 1," the command bar named "Custom" is displayed. If User is any other value, the built-in Visual Basic command bar is reset to its default state and the command bar named "Custom" is disabled.


```
Set myBar = CommandBars("Custom") 
If user = "Level 1" Then 
    myBar.Visible =  True 
Else 
    CommandBars("Visual Basic").Reset 
    myBar.Enabled = False  
End If
```


## See also


#### Concepts


[CommandBar Object](commandbar-object-office.md)
#### Other resources


[CommandBar Object Members](commandbar-members-office.md)

