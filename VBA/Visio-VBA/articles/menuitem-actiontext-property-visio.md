---
title: MenuItem.ActionText Property (Visio)
keywords: vis_sdr.chm12913015
f1_keywords:
- vis_sdr.chm12913015
ms.prod: visio
api_name:
- Visio.MenuItem.ActionText
ms.assetid: 293d60d4-11fd-52f7-c934-3cc56a632659
ms.date: 06/08/2017
---


# MenuItem.ActionText Property (Visio)

Gets or sets the action text for a menu item. Read/write.


## Syntax

 _expression_ . **ActionText**

 _expression_ A variable that represents a **MenuItem** object.


### Return Value

String


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

Action text is a string that describes the action on the  **Undo**,  **Redo**, and  **Repeat** menu items on the **Quick Access** toolbar.

If the  **ActionText** property is empty and the object's **CmdNum** property is set to one of the Microsoft Visio built-in command IDs, the item uses the default action text from the built-in Visio user interface.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to set a menu item's  **ActionText** property. It also shows how to add a menu and menu item to the drawing window menu set. This example assumes that you already have a macro in the current Visual Basic project. Before running this macro, replace _macroname_ with the name of your macro.

 Beginning with Microsoft Visio 2002, the **AddOnName** property used in this example cannot execute a string that contains arbitrary Microsoft Visual Basic code. To call code that in previous versions of Visio you would have passed to the **AddOnName** property, move it to a procedure in a document's Visual Basic project that is called from the **AddOnName** property, as shown in this example.

To restore the built-in user interface in Microsoft Visio after you run this macro, call the  **ThisDocument.ClearCustomMenus** method.




```vb
 
Public Sub ActionText_Example() 
 
 Dim vsoUIObject As Visio.UIObject 
 Dim vsoMenuSets As Visio.MenuSets 
 Dim vsoMenuSet As Visio.MenuSet 
 Dim vsoMenus As Visio.Menus 
 Dim vsoMenu As Visio.Menu 
 Dim vsoMenuItems As Visio.MenuItems 
 Dim vsoMenuItem As Visio.MenuItem 
 
 'Get a UIObject object that represents Visio built-in menus 
 Set vsoUIObject = Visio.Application.BuiltInMenus 
 
 'Get the MenuSets collection 
 Set vsoMenuSets = vsoUIObject.MenuSets 
 
 'Get the drawing window menu set 
 Set vsoMenuSet = vsoMenuSets.ItemAtID(visUIObjSetDrawing) 
 
 'Get the Menus collection. 
 Set vsoMenus = vsoMenuSet.Menus 
 
 'Add a Demo menu before the Window menu 
 Set vsoMenu = vsoMenus.AddAt(7) 
 vsoMenu.Caption = "Demo" 
 
 'Get the MenuItems collection 
 Set vsoMenuItems = vsoMenu.MenuItems 
 
 'Add a menu item to the new Demo menu 
 Set vsoMenuItem = vsoMenuItems.Add 
 
 'Set the properties for the new menu item 
 vsoMenuItem.Caption = "&;macroname " 
 vsoMenuItem.AddOnName = "ThisDocument.macroname " 
 vsoMenuItem.AddOnArgs = "/Arg1 = True" 
 vsoMenuItem.ActionText = "Run(macroname )" 
 
 'Tell Visio to use the new UI when the document is active 
 ThisDocument.SetCustomMenus vsoUIObject 
 
End Sub
```


