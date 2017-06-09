---
title: CommandBar.Type Property (Office)
keywords: vbaof11.chm3019
f1_keywords:
- vbaof11.chm3019
ms.prod: office
api_name:
- Office.CommandBar.Type
ms.assetid: e023edd9-a8f4-c20f-c6b1-c434182bd748
ms.date: 06/08/2017
---


# CommandBar.Type Property (Office)

Gets the type of command bar. Read-only.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **Type**

 _expression_ Required. A variable that represents a **[CommandBar](commandbar-object-office.md)** object.


## Example

This example finds the first control on the command bar named "Custom." Using the  **Type** property, the example determines whether the control is a button. If the control is a button, the example copies the face of the **Copy** button (on the **Standard** toolbar) and then pastes it onto the control.


```
Set oldCtrl = CommandBars("Custom").Controls(1) 
If oldCtrl.Type = msoControlButton Then 
    Set newCtrl = CommandBars.FindControl(Type:= _ 
        MsoControlButton, ID:= _ 
        CommandBars("Standard").Controls("Copy").ID) 
    NewCtrl.CopyFace 
    OldCtrl.PasteFace 
End If
```


## See also


#### Concepts


[CommandBar Object](commandbar-object-office.md)
#### Other resources


[CommandBar Object Members](commandbar-members-office.md)

