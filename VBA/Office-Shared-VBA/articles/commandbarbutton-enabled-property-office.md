---
title: CommandBarButton.Enabled Property (Office)
ms.prod: office
api_name:
- Office.CommandBarButton.Enabled
ms.assetid: 264335ca-6506-0e86-16df-44af277ade83
ms.date: 06/08/2017
---


# CommandBarButton.Enabled Property (Office)

 **True** if the specified **CommandBar** or **CommandBarControl** is enabled. Read/write .


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **Enabled**

 _expression_ A variable that represents a **[CommandBarButton](commandbarbutton-object-office.md)** object.


### Return Value

Boolean


## Remarks

For command bars, setting this property to  **True** causes the name of the command bar to appear in the list of available command bars.

For built-in controls, if you set the  **Enabled** property to **True**, the application determines its state, but setting it to **False** will force it to be disabled.


## See also


#### Concepts


[CommandBarButton Object](commandbarbutton-object-office.md)
#### Other resources


[CommandBarButton Object Members](commandbarbutton-members-office.md)

