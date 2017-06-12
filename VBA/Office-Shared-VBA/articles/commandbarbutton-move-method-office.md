---
title: CommandBarButton.Move Method (Office)
ms.prod: office
api_name:
- Office.CommandBarButton.Move
ms.assetid: b2d462ec-63a7-a395-8d93-bedbf1d6941d
ms.date: 06/08/2017
---


# CommandBarButton.Move Method (Office)

Moves the specified  **CommandBarButton** control to an existing command bar.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **Move**( **_Bar_**, **_Before_** )

 _expression_ Required. A variable that represents a **[CommandBarButton](commandbarbutton-object-office.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Bar_|Optional|**Variant**|A  **Command** object that represents the destination command bar for the control. If this argument is omitted, the control is moved to the end of the command bar where the control currently resides.|
| _Before_|Optional|**Variant**|A number that indicates the position for the control. The control is inserted before the control currently occupying this position. If this argument is omitted, the control is inserted on the same command bar.|

## Example

This example moves the first combo box control on the command bar named Custom to the position before the seventh control on that command bar. The example sets the tag to "Selection box" and assigns the control a low priority so that it will likely be dropped from the command bar if all the controls don't fit in one row.


```
Set allcontrols = CommandBars("Custom").Controls 
For Each ctrl In allControls 
    If ctrl.Type = msoControlComboBox Then 
        With ctrl 
            .Move Before:=7 
             .Tag = "Selection box" 
             .Priority = 5 
         End With 
         Exit For 
    End If 
Next
```


## See also


#### Concepts


[CommandBarButton Object](commandbarbutton-object-office.md)
#### Other resources


[CommandBarButton Object Members](commandbarbutton-members-office.md)

