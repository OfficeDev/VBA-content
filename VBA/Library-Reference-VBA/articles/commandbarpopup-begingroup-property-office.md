---
title: CommandBarPopup.BeginGroup Property (Office)
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.CommandBarPopup.BeginGroup
ms.assetid: 0ecc5c98-5db7-792c-8f33-86f7df32d912
---


# CommandBarPopup.BeginGroup Property (Office)

Gets  **True** if the specified command bar control appears at the beginning of a group of controls on the command bar. Read/write.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **BeginGroup**

 _expression_ A variable that represents a **CommandBarPopup** object.


### Return Value

Boolean


## Example

This example begins a new group with the last control on the active menu bar.


```vb
Set myMenuBar = CommandBars.ActiveMenuBar 
Set lastMenu = myMenuBar _ 
    .Controls(myMenuBar.Controls.Count) 
lastMenu.BeginGroup = True
```


## See also


#### Concepts


[CommandBarPopup Object](commandbarpopup-object-office.md)

