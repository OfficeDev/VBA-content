---
title: CommandBarControl.Enabled Property (Office)
ms.prod: office
api_name:
- Office.CommandBarControl.Enabled
ms.assetid: 74105bf5-96a0-09ea-bb00-ef102705372c
ms.date: 06/08/2017
---


# CommandBarControl.Enabled Property (Office)

Gets or sets a  **Boolean** value specifying if the **CommandBarControl** is enabled. Read/write.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **Enabled**

 _expression_ A variable that represents a **[CommandBarControl](commandbarcontrol-object-office.md)** object.


## Remarks

For command bars, setting this property to  **True** causes the name of the command bar to appear in the list of available command bars.

For built-in controls, if you set the  **Enabled** property to **True**, the application determines its state, but setting it to **False** will force it to be disabled.


## See also


#### Concepts


[CommandBarControl Object](commandbarcontrol-object-office.md)
#### Other resources


[CommandBarControl Object Members](commandbarcontrol-members-office.md)

