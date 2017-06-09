---
title: CommandBarControl.Priority Property (Office)
ms.prod: office
api_name:
- Office.CommandBarControl.Priority
ms.assetid: 1bb78346-a815-75f8-f2f6-8ecff2b54cbd
ms.date: 06/08/2017
---


# CommandBarControl.Priority Property (Office)

Gets or sets the priority of a  **CommandBarControl**. A control's priority determines whether the control can be dropped from a docked command bar if the command bar controls can't fit in a single row. Controls that can't fit in a single row drop off command bars from right to left. Read/write.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **Priority**

 _expression_ A variable that represents a **[CommandBarControl](commandbarcontrol-object-office.md)** object.


## Remarks

Valid priority numbers are 0 (zero) through 7 and the default value is 3. A priority of 1 means that the control cannot be dropped from a toolbar. Other priority values are ignored.

The  **Priority** property is not used by command bar controls that are menu items.


## See also


#### Concepts


[CommandBarControl Object](commandbarcontrol-object-office.md)
#### Other resources


[CommandBarControl Object Members](commandbarcontrol-members-office.md)

