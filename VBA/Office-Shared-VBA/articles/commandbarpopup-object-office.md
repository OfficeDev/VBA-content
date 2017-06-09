---
title: CommandBarPopup Object (Office)
keywords: vbaof11.chm7000
f1_keywords:
- vbaof11.chm7000
ms.prod: office
api_name:
- Office.CommandBarPopup
ms.assetid: a8ae06a3-1d7b-a531-91df-756fafee5314
ms.date: 06/08/2017
---


# CommandBarPopup Object (Office)

Represents a pop-up control on a command bar.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Remarks

Every pop-up control contains a  **CommandBar** object. To return the command bar from a pop-up control, apply the **CommandBar** property to the **CommandBarPopup** object.

 Use Controls(index), where _index_ is the number of the control, to return a **CommandBarPopup** object. Note that the **Type** property of the control must be **msoControlPopup**, **msoControlGraphicPopup**, **msoControlButtonPopup**, **msoControlSplitButtonPopup**, or **msoControlSplitButtonMRUPopup**.


## Example

You can also use the  **FindControl** method to return a **CommandBarPopup** object. The following example searches all command bars for a **CommandBarPopup** object whose tag is "Graphics."


```
Set myControl = Application.CommandBars.FindControl _ 
(Type:=msoControlPopup, Tag:="Graphics")
```


## See also


#### Concepts


[Object Model Reference](reference-object-library-reference-for-office.md)
#### Other resources


[CommandBarPopup Object Members](commandbarpopup-members-office.md)

