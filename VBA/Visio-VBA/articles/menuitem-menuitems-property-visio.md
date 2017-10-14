---
title: MenuItem.MenuItems Property (Visio)
keywords: vis_sdr.chm12913905
f1_keywords:
- vis_sdr.chm12913905
ms.prod: visio
api_name:
- Visio.MenuItem.MenuItems
ms.assetid: eee46c53-4b9e-806d-c0be-7a4f907ee314
ms.date: 06/08/2017
---


# MenuItem.MenuItems Property (Visio)

Returns the  **MenuItems** collection of a **MenuItem** object. Read-only.


## Syntax

 _expression_ . **MenuItems**

 _expression_ A variable that represents a **MenuItem** object.


### Return Value

MenuItems


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

If a  **Menu** object represents a hierarchical menu, its **MenuItems** collection contains submenu items. Otherwise, its **MenuItems** collection is empty.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Menus** property to get the **Menus** collection of a **MenuSet** object. It adds a menu and menu item to the user interface and sets the **Caption** property of the menu and menu item.

To restore the built-in user interface in Microsoft Visio after you run this macro, call the  **ThisDocument.ClearCustomMenus** method.




```vb
 
Public Sub Menus_Example() 
 
 Dim vsoUIObject As Visio.UIObject 
 Dim vsoMenuSets As Visio.MenuSets 
 Dim vsoMenuSet As Visio.MenuSet 
 Dim vsoMenus As Visio.Menus 
 Dim vsoMenu As Visio.Menu 
 Dim vsoMenuItems As Visio.MenuItems 
 Dim vsoMenuItem As Visio.MenuItem 
 
 'Get a UIObject object that represents Microsoft Office Visio built-in menus. 
 Set vsoUIObject = Visio.Application.BuiltInMenus 
 
 'Get the MenuSets collection. 
 Set vsoMenuSets = vsoUIObject.MenuSets 
 
 'Get the drawing window menu set. 
 Set vsoMenuSet = vsoMenuSets.ItemAtId(visUIObjSetDrawing) 
 
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


