---
title: CommandBarComboBox.Copy Method (Office)
ms.prod: office
api_name:
- Office.CommandBarComboBox.Copy
ms.assetid: 15eb757c-bb07-cd98-ff9e-1810db4f475c
ms.date: 06/08/2017
---


# CommandBarComboBox.Copy Method (Office)

Copies a command bar combo box control to an existing command bar.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **Copy**( **_Bar_**, **_Before_** )

 _expression_ A variable that represents a **CommandBarComboBox** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Bar_|Optional|**Variant**| A **CommandBar** object that represents the destination command bar. If this argument is omitted, the control is copied to the command bar where the control already exists.|
| _Before_|Optional|**Variant**|A number that indicates the position for the new control on the command bar. The new control will be inserted before the control at this position. If this argument is omitted, the control is copied to the end of the command bar.|

### Return Value

CommandBarControl


## See also


#### Concepts


[CommandBarComboBox Object](commandbarcombobox-object-office.md)
#### Other resources


[CommandBarComboBox Object Members](commandbarcombobox-members-office.md)

