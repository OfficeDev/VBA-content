---
title: CommandBarPopup.Priority Property (Office)
ms.prod: office
api_name:
- Office.CommandBarPopup.Priority
ms.assetid: cef115fd-fdc8-d8a3-b51d-c9fbc21a810f
ms.date: 06/08/2017
---


# CommandBarPopup.Priority Property (Office)

Gets or sets the priority of a  **CommandBarPopup** control. Read/write.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **Priority**

 _expression_ A variable that represents a **CommandBarPopup** object.


### Return Value

Integer


## Remarks

 A control's priority determines whether the control can be dropped from a docked command bar if the command bar controls can't fit in a single row. Controls that can't fit in a single row drop off command bars from right to left.


## Example

The following example sets the descriptive text and priority of a command bar popup.


```
Dim popControl As CommandBarPopup 
Set popControl = Application.CommandBars.FindControl _ 
(Type:=msoControlPopup, Tag:="Graphics") 
 
With popControl. 
            .DescriptionText = "Graphics Selection dialog" 
            .Priority = 5 
End With 

```


## See also


#### Concepts


[CommandBarPopup Object](commandbarpopup-object-office.md)
#### Other resources


[CommandBarPopup Object Members](commandbarpopup-members-office.md)

