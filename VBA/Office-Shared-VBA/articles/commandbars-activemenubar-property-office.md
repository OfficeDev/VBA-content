---
title: CommandBars.ActiveMenuBar Property (Office)
keywords: vbaof11.chm2002
f1_keywords:
- vbaof11.chm2002
ms.prod: office
api_name:
- Office.CommandBars.ActiveMenuBar
ms.assetid: 8f341f53-418c-6d05-ac0b-e45a6b2baa0d
ms.date: 06/08/2017
---


# CommandBars.ActiveMenuBar Property (Office)

Gets a  **CommandBar** object that represents the active menu bar in the container application. Read-only.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **ActiveMenuBar**

 _expression_ A variable that represents a **CommandBars** object.


## Example

This example adds a temporary pop-up control named "Custom" to the end of the active menu bar, and adds a control named "Import" to the pop-up control.


```
Set myMenuBar = CommandBars.ActiveMenuBar 
Set newMenu = myMenuBar.Controls.Add(Type:=msoControlPopup, Temporary:=True) 
newMenu.Caption = "Custom" 
Set ctrl1 = newMenu.CommandBar.Controls _ 
    .Add(Type:=msoControlButton, Id:=1) 
With ctrl1 
    .Caption = "Import" 
    .TooltipText = "Import" 
    .Style = msoButtonCaption 
End With
```


## See also


#### Concepts


[CommandBars Object](commandbars-object-office.md)
#### Other resources


[CommandBars Object Members](commandbars-members-office.md)

