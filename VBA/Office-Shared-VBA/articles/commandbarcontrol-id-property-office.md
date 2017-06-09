---
title: CommandBarControl.Id Property (Office)
ms.prod: office
api_name:
- Office.CommandBarControl.Id
ms.assetid: 0931a07a-4a6b-cc84-a43b-b57ea9a22b78
ms.date: 06/08/2017
---


# CommandBarControl.Id Property (Office)

Gets the ID for a built-in  **CommandBarControl**. Read-only.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **Id**

 _expression_ Required. A variable that represents a **[CommandBarControl](commandbarcontrol-object-office.md)** object.


## Remarks

A control's ID determines the built-in action for that control. The value of the  **Id** property for all custom controls is 1.


## Example

This example changes the button face of the first control on the command bar named "Custom2" if the button's  **Id** value is less than 25.


```
Set ctrl = CommandBars("Custom").Controls(1) 
With ctrl 
    If .Id < 25 Then 
        .FaceId = 17 
        .Tag = "Changed control" 
    End If 
End With
```

The following example changes the caption of every control on the toolbar named "Standard" to the current value of the  **Id** property for that control.




```
For Each ctl In CommandBars("Standard").Controls 
    ctl.Caption = CStr(ctl.Id) 
Next ctl
```


## See also


#### Concepts


[CommandBarControl Object](commandbarcontrol-object-office.md)
#### Other resources


[CommandBarControl Object Members](commandbarcontrol-members-office.md)

