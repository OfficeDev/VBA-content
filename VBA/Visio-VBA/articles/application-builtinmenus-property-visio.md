---
title: Application.BuiltInMenus Property (Visio)
keywords: vis_sdr.chm10013160
f1_keywords:
- vis_sdr.chm10013160
ms.prod: visio
api_name:
- Visio.Application.BuiltInMenus
ms.assetid: 0f76537c-5d9b-bcfa-c528-4644bd0375d5
ms.date: 06/08/2017
---


# Application.BuiltInMenus Property (Visio)

Returns a  **UIObject** object that represents a copy of the built-in Microsoft Visio menus and accelerators. Read-only.


## Syntax

 _expression_ . **BuiltInMenus**

 _expression_ A variable that represents an **Application** object.


### Return Value

UIObject


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

You can use the  **BuiltInMenus** property to obtain a **UIObject** object and modify its menus and accelerators. You can then use the **SetCustomMenus** method of an **Application** or **Document** object to add your customized menus and accelerators to the built-in Visio user interface.

You can also use the  **SaveToFile** method of the **UIObject** object to store its menus in a file and reload them as custom menus by setting the **CustomMenusFile** property of an **Application** or **Document** object.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **BuiltInMenus** property. It adds a menu and menu item to the **Add-ins** tab and sets the menu and menu item's **Caption** properties.

To restore the built-in user interface in Microsoft Visio after you run this macro, call the  **ThisDocument.ClearCustomMenus** method.




```vb
 
Public Sub BuiltInMenus_Example() 
 
 Dim vsoUIObject As Visio.UIObject 
 Dim vsoMenuSets As Visio.MenuSets 
 Dim vsoMenuSet As Visio.MenuSet 
 Dim vsoMenus As Visio.Menus 
 Dim vsoMenu As Visio.Menu 
 Dim vsoMenuItems As Visio.MenuItems 
 Dim vsoMenuItem As Visio.MenuItem 
 
 'Get a UIObject object that represents Visio built-in menus. 
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


