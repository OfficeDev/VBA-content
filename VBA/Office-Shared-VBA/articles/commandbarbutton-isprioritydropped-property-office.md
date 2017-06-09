---
title: CommandBarButton.IsPriorityDropped Property (Office)
ms.prod: office
api_name:
- Office.CommandBarButton.IsPriorityDropped
ms.assetid: 68398973-675f-2180-b22c-4ad5de0582f7
ms.date: 06/08/2017
---


# CommandBarButton.IsPriorityDropped Property (Office)

Gets  **True** if the **CommandBarButton** control is currently dropped from the menu or toolbar based on usage statistics and layout space. (Note that this is not the same as the control's visibility, as set by the Visible property). Read-only.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **IsPriorityDropped**

 _expression_ A variable that represents a **CommandBarButton** object.


### Return Value

Boolean


## Remarks

 A control with **Visible** set to **True**, will not be immediately visible on a personalized menu or toolbar if **IsPriorityDropped** is **True**.

To determine when to set  **IsPriorityDropped** to **True** for a specific menu item, Microsoft Office maintains a total count of the number of times the menu item was used and a record of the number of different application sessions in which the user has used another menu item in the same menu as this menu item, without using the specific menu item. When this value reaches certain threshold values, the count is decremented. When the count reaches zero, **IsPriorityDropped** is set to **True**. Programmers cannot set the session value, the threshold value, or the **IsPriorityDropped** property. Programmers can, however, use the **AdaptiveMenus** property to disable adaptive menus for specific menus in an application.

To determine when to set  **IsPriorityDropped** to **True** for a specific toolbar control, Office maintains a list of the order in which all the controls on that toolbar were last executed. A toolbar will always show as many controls as it has space to show, in the order of most recently used to least recently used. Controls with **Priority** set to 1 will always be shown and the toolbar will wrap rows, if necessary, to show these controls. Programmers can use the **Priority** property to ensure that specific toolbar controls are always shown, or to reposition toolbars so that they have enough space to display all of their controls.

You can use the following table to predict the number of sessions for which a menu item on a personalized menu will remain visible before the menu item's  **IsPriorityDropped** property is set to **True**.



|**Number of uses of the command bar control**|**Number of sessions of the application**|
|:-----|:-----|
|0, 1|3|
|2|6|
|3|9|
|4, 5|12|
|6- 8|17|
|9-13|23|
|14-24|29|
|25 or more|31|

## See also


#### Concepts


[CommandBarButton Object](commandbarbutton-object-office.md)
#### Other resources


[CommandBarButton Object Members](commandbarbutton-members-office.md)

