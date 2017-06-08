---
title: Menu.Caption Property (Visio)
keywords: vis_sdr.chm13113170
f1_keywords:
- vis_sdr.chm13113170
ms.prod: visio
api_name:
- Visio.Menu.Caption
ms.assetid: e89db3a6-59cc-a87a-dfbd-f8c18e7c4c58
ms.date: 06/08/2017
---


# Menu.Caption Property (Visio)

Gets or sets the caption for a menu. Read/write.


## Syntax

 _expression_ . **Caption**

 _expression_ A variable that represents a **Menu** object.


### Return Value

String


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.




- Use an ampersand (&;) in the string to cause the next character in the string to become the shortcut key for that menu or menu item. For example, the string "F _&;o_ rmat" causes _o_ to become the shortcut key for that menu item in that one menu.
    
- Use "" in the string to display a double quotation mark on the menu.
    
- Use &;&; in the string to display an ampersand on the menu.
    



## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Caption** property. It adds a menu and menu item to the **Add-ins** tab and sets the menu and menu item's **Caption** properties.

To restore the built-in user interface in Microsoft Visio after you run this macro, call the  **ThisDocument.ClearCustomMenus** method.




```vb
 
Public Sub Caption_Example() 
 
 Dim vsoUIObject As Visio.UIObject 
 Dim vsoMenuSets As Visio.MenuSets 
 Dim vsoMenuSet As Visio.MenuSet 
 Dim vsoMenus As Visio.Menus 
 Dim vsoMenu As Visio.Menu 
 Dim vsoMenuItems As Visio.MenuItems 
 Dim vsoMenuItem As Visio.MenuItem 
 
 'Get a UIObject object that represents Microsoft Visio built-in menus. 
 Set vsoUIObject = Visio.Application.BuiltInMenus 
 
 'Get the MenuSets collection. 
 Set vsoMenuSets = vsoUIObject.MenuSets 
 
 'Get the drawing window menu set. 
 Set vsoMenuSet = vsoMenuSets.ItemAtID(visUIObjSetDrawing) 
 
 'Get the Menus collection. 
 Set vsoMenus = vsoMenuSet.Menus 
 
 'Add a new menu before the Window menu. 
 Set vsoMenu = vsoMenus.AddAt(7) 
 vsoMenu.Caption = "MyNewMenu" 
 
 'Get the MenuItems collection. 
 Set vsoMenuItems = vsoMenu.MenuItems 
 
 'Add a menu item to the new menu. 
 Set vsoMenuItem = vsoMenuItems.Add 
 
 'Set the Caption property for the new menu item. 
 vsoMenuItem.Caption = "&;MyNewMenuItem" 
 
 'Tell Visio to use the new UI when the document is active. 
 ThisDocument.SetCustomMenus vsoUIObject 
 
End Sub
```


