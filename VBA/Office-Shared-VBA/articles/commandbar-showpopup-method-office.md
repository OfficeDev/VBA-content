---
title: CommandBar.ShowPopup Method (Office)
keywords: vbaof11.chm3017
f1_keywords:
- vbaof11.chm3017
ms.prod: office
api_name:
- Office.CommandBar.ShowPopup
ms.assetid: e501b7d2-2606-976c-b391-1aa8fa07f105
ms.date: 06/08/2017
---


# CommandBar.ShowPopup Method (Office)

Displays a command bar as a shortcut menu at the specified coordinates or at the current pointer coordinates.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **ShowPopup**( **_x_**, **_y_** )

 _expression_ A variable that represents a **CommandBar** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _x_|Optional|**Variant**|The x-coordinate on which the location of the shortcut menu is based. If this argument is omitted, the current x-coordinate of the pointer is used.|
| _y_|Optional|**Variant**|The y-coordinate on which the location of the shortcut menu is based. If this argument is omitted, the current y-coordinate of the pointer is used.|

## Remarks

When menus are left-aligned, the shortcut menu displayed by the  **ShowPopup** method has its upper left corner at (x, y + 1); when menus are right-aligned, the shortcut menu has its upper right corner at (x + 1, y + 1). You can use the Windows function **GetSystemMetrics(SM_MENUDROPALIGNMENT)** to check the system metric for dropdown menu alignment.

When the screen location of the (x, y) coordinates would cause all or part of the popup menu to be displayed beyond the edge of the visible screen, then the popup menu shifts to fit into the viewable area.


## Example

This example creates a shortcut menu containing two controls. The  **ShowPopup** method is used to make the shortcut menu visible.


```
Set myBar = CommandBars _ 
    .Add(Name:="Custom", Position:=msoBarPopup, Temporary:=False) 
With myBar 
    .Controls.Add Type:=msoControlButton, Id:=3 
    .Controls.Add Type:=msoControlComboBox 
End With 
myBar.ShowPopup
```


 **Note**  


 **Note**  If the  **Position** property of the command bar is not set to **msoBarPopup**, this method fails.


## See also


#### Concepts


[CommandBar Object](commandbar-object-office.md)
#### Other resources


[CommandBar Object Members](commandbar-members-office.md)

