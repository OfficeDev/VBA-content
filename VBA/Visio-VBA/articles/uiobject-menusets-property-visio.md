---
title: UIObject.MenuSets Property (Visio)
keywords: vis_sdr.chm14913915
f1_keywords:
- vis_sdr.chm14913915
ms.prod: visio
api_name:
- Visio.UIObject.MenuSets
ms.assetid: 8acecfc4-5a49-e11f-b9e9-07d5a464681a
ms.date: 06/08/2017
---


# UIObject.MenuSets Property (Visio)

Returns the  **MenuSets** collection of a **UIObject** object. Read-only.


## Syntax

 _expression_ . **MenuSets**

 _expression_ A variable that represents a **UIObject** object.


### Return Value

MenuSets


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

If a  **UIObject** object represents menus and accelerators (for example, if the object was retrieved using the **BuiltInMenus** property of an **Application** or **Document** object), its **MenuSets** collection represents all of the menus for that **UIObject** object.

Use the  **ItemAtID** property of a **MenuSets** object to retrieve menus for a particular window context such as the drawing window. If a context does not include menus, it has no **MenuSets** collection.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **MenuSets** property to get the **MenuSets** collection of a **UIObject** object. It adds a menu and menu item to the drawing window menu set and sets the menu and menu item's **Caption** properties.

To restore the built-in menus in Microsoft Visio after you run this macro, call the  **ThisDocument.ClearCustomMenus** method.




```vb
Public Sub MenuSets_Example() 
 
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


