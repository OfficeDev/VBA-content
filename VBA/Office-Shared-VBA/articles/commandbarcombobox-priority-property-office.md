---
title: CommandBarComboBox.Priority Property (Office)
ms.prod: office
api_name:
- Office.CommandBarComboBox.Priority
ms.assetid: 0166df8f-316a-8414-a3af-1156fc1a1166
ms.date: 06/08/2017
---


# CommandBarComboBox.Priority Property (Office)

Gets or sets the priority of a  **CommandBarComboBox** control. A control's priority determines whether the control can be dropped from a docked command bar if the command bar controls can't fit in a single row. Read/write.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **Priority**

 _expression_ A variable that represents a **[CommandBarComboBox](commandbarcombobox-object-office.md)** object.


## Remarks

Valid priority numbers are 0 (zero) through 7 and the default value is 3. A priority of 1 means that the control cannot be dropped from a toolbar. Other priority values are ignored.

The  **Priority** property is not used by command bar controls that are menu items.


## See also


#### Concepts


[CommandBarComboBox Object](commandbarcombobox-object-office.md)
#### Other resources


[CommandBarComboBox Object Members](commandbarcombobox-members-office.md)

