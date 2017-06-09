---
title: CommandBarComboBox Object (Office)
keywords: vbaof11.chm243000
f1_keywords:
- vbaof11.chm243000
ms.prod: office
api_name:
- Office.CommandBarComboBox
ms.assetid: fcfe6bde-dea0-f1f1-ad30-d0e28f97dd07
ms.date: 06/08/2017
---


# CommandBarComboBox Object (Office)

Represents a combo box control on a command bar.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Remarks

Use  **Controls(index)**, where _index_ is the index number of the control, to return a **CommandBarComboBox** object. Note that the **Type** property of the control must be **msoControlEdit**, **msoControlDropdown**, **msoControlComboBox**, **msoControlButtonDropdown**, **msoControlSplitDropdown**, **msoControlOCXDropdown**, **msoControlGraphicCombo**, or **msoControlGraphicDropdown**.


## Example

The following example adds two items to the second control on the command bar named  **Custom**, and then it adjusts the size of the control.


```
Set combo = CommandBars("Custom").Controls(2) 
With combo 
    .AddItem "First Item", 1 
    .AddItem "Second Item", 2 
    .DropDownLines = 3 
    .DropDownWidth = 75 
    .ListIndex = 0 
End With
```

You can also use the  **FindControl** method to return a **CommandBarComboBox** object. The following example searches all command bars for a visible **CommandBarComboBox** object whose tag is "sheet assignments."




```
Set myControl = CommandBars.FindControl _ 
(Type:=msoControlComboBox, Tag:="sheet assignments", Visible:=True)
```


## See also


#### Concepts


[Object Model Reference](reference-object-library-reference-for-office.md)
#### Other resources


[CommandBarComboBox Object Members](commandbarcombobox-members-office.md)

