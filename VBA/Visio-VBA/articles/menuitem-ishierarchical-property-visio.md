---
title: MenuItem.IsHierarchical Property (Visio)
keywords: vis_sdr.chm12913740
f1_keywords:
- vis_sdr.chm12913740
ms.prod: visio
api_name:
- Visio.MenuItem.IsHierarchical
ms.assetid: d8643162-6d8a-4558-d4e0-c563af680cb3
ms.date: 06/08/2017
---


# MenuItem.IsHierarchical Property (Visio)

Indicates whether a menu item is hierarchical; that is, whether it contains a drop-down menu that contains more items, which can in turn be accessed by iterating through the  **MenuItems** collection of the menu item. Read-only.


## Syntax

 _expression_ . **IsHierarchical**

 _expression_ A variable that represents a **MenuItem** object.


### Return Value

Integer


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

The value of the  **CmdNum** property of a **MenuItem** object that represents a hierarchical menu should be zero (0). This corresponds to the Microsoft Visio constant **visCmdHierarchical** .


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **IsHierarchical** property to delete a hierarchical menu.

To restore Visio's built-in menus after you run this macro, call the  **ThisDocument.ClearCustomMenus** method.




```vb
 
Public Sub IsHierarchical_Example() 
 
 Dim vsoUIObject As Visio.UIObject 
 Dim vsoMenuSet As Visio.MenuSet 
 Dim vsoMenu As Visio.Menu 
 Dim vsoMenuItems As Visio.MenuItems 
 Dim vsoMenuItem As Visio.MenuItem 
 Dim vsoHierarchicalMenuItems As Visio.MenuItems 
 Dim vsoHierarchicalMenuItem As Visio.MenuItem 
 
 'True if variable represents a hierarchical menu item. 
 Dim blsHierarchicalState As Boolean 
 
 Dim intCounterOuter As Integer 
 Dim intCounterInner As Integer 
 
 'Get the UIOject object for the copy of the built-in menus. 
 Set vsoUIObject = Visio.Application.BuiltInMenus 
 
 'Set vsoMenuSet to the drawing menu set. 
 Set vsoMenuSet = vsoUIObject.MenuSets.ItemAtID(visUIObjSetDrawing) 
 
 'Get the Tools menu. Because you got the built-in 
 'menus, you know that you can find the Tools menu by its 
 'position. If you had retrieved a custom UI, you would have 
 'to loop through the menus checking the caption to find 
 'the Tools menu. When you use a custom menu, there is no 
 'guarantee that you will find a Tools menu, because it 
 'could have been deleted. 
 Set vsoMenu = vsoMenuSet.Menus(5) 
 
 'Get the MenuItems collection for the Tools menu. 
 Set vsoMenuItems = vsoMenu.MenuItems 
 
 'Locate the Macros menu item. Because you got the 
 'built-in menus, you know you will find it. If you had 
 'started from a custom menu, you would need to handle 
 'the case of not finding the menu item. 
 For intCounterOuter = 0 To vsoMenuItems.Count - 1 
 
 'Retrieve the current menu item from the collection. 
 Set vsoMenuItem = vsoMenuItems(intCounterOuter) 
 
 'Check CmdNum to see if it is Macro. 
 If vsoMenuItem.CmdNum = visCmdHierarchical And _ 
 vsoMenuItem.Caption = "&;Macros" Then 
 
 'The value of blsHierarchicalState is True. 
 blsHierarchicalState = vsoMenuItem.IsHierarchical 
 
 'Get the MenuItems collection for the 
 'hierarchical menu. 
 Set vsoHierarchicalMenuItems = vsoMenuItem.MenuItems 
 
 'Locate the Visual Basic Editor menu item. 
 'As with the Macros menu item, you know you will 
 'find the Visual Basic Editor menu item 
 'because you started with a copy of 
 'the built-in menus. 
 For intCounterInner = 0 To vsoHierarchicalMenuItems.Count - 1 
 
 'Retrieve menu item from collection. 
 Set vsoHierarchicalMenuItem = vsoHierarchicalMenuItems(intCounterInner) 
 
 'Check CmdNum. 
 If vsoHierarchicalMenuItem.CmdNum = visCmdToolsRunVBE Then 
 
 'Delete the Visual Basic Editor menu item. 
 vsoHierarchicalMenuItem.Delete 
 
 'Exit the inner For loop. 
 Exit For 
 
 End If 
 
 Next intCounterInner 
 
 'Exit the outer For loop. 
 Exit For 
 
 End If 
 
 Next intCounterOuter 
 
 'Tell Microsoft Visio to use the custom user interface 'while the document is active. 
 ThisDocument.SetCustomMenus vsoUIObject 
 
End Sub
```


