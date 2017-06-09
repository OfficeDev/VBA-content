---
title: CommandBarControls Object (Office)
keywords: vbaof11.chm4000
f1_keywords:
- vbaof11.chm4000
ms.prod: office
api_name:
- Office.CommandBarControls
ms.assetid: 7ccae243-2870-95c2-1e08-140a3e638fe6
ms.date: 06/08/2017
---


# CommandBarControls Object (Office)

A collection of  **CommandBarControl** objects that represent the command bar controls on a command bar.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Example

Use the  **Controls** property to return the **CommandBarControls** collection. The following example changes the caption of every control on the toolbar named "Standard" to the current value of the **Id** property for that control.


```
For Each ctl In CommandBars("Standard").Controls 
    ctl.Caption = CStr(ctl.Id) 
Next ctl
```

Use the  **Add** method to add a new command bar control to the **CommandBarControls** collection. This example adds a new, blank button to the command bar named "Custom."




```
Set myBlankBtn = CommandBars("Custom").Controls.Add
```

Use Controls(index), where  _index_ is the caption or index number of a control, to return a **CommandBarControl**, **CommandBarButton**, **CommandBarComboBox**, or **CommandBarPopup** object. The following example copies the first control from the command bar named "Standard" to the command bar named "Custom."




```
Set myCustomBar = CommandBars("Custom") 
Set myControl = CommandBars("Standard").Controls(1) 
myControl.Copy Bar:=myCustomBar, Before:=1
```


## See also


#### Concepts


[Object Model Reference](reference-object-library-reference-for-office.md)
#### Other resources


[CommandBarControls Object Members](commandbarcontrols-members-office.md)

