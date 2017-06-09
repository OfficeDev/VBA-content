---
title: CommandBarPopup.CommandBar Property (Office)
keywords: vbaof11.chm7001
f1_keywords:
- vbaof11.chm7001
ms.prod: office
api_name:
- Office.CommandBarPopup.CommandBar
ms.assetid: e78abe18-d260-8cac-d647-322b449e4bbb
ms.date: 06/08/2017
---


# CommandBarPopup.CommandBar Property (Office)

Gets a  **CommandBar** object that represents the menu displayed by the specified pop-up control. Read-only.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **CommandBar**

 _expression_ A variable that represents a **CommandBarPopup** object.


## Example

This example sets the variable fourthLevel to the fourth control on the command bar named "Drawing."


```
Set fourthLevel = CommandBars("Drawing") _ 
    .Controls(1).CommandBar.Controls(4)
```


## See also


#### Concepts


[CommandBarPopup Object](commandbarpopup-object-office.md)
#### Other resources


[CommandBarPopup Object Members](commandbarpopup-members-office.md)

