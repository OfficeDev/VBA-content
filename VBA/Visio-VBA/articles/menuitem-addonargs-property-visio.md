---
title: MenuItem.AddOnArgs Property (Visio)
keywords: vis_sdr.chm12913045
f1_keywords:
- vis_sdr.chm12913045
ms.prod: visio
api_name:
- Visio.MenuItem.AddOnArgs
ms.assetid: 71e4be8f-3176-c3e8-c25f-7d58efef9ab6
ms.date: 06/08/2017
---


# MenuItem.AddOnArgs Property (Visio)

Gets or sets the argument string that you send to the add-on associated with a particular menu item. Read/write.


## Syntax

 _expression_ . **AddOnArgs**

 _expression_ A variable that represents a **MenuItem** object.


### Return Value

String


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

An argument's string can be anything appropriate for the add-on. However, the arguments are packaged together with other information into a command string, which cannot exceed 127 characters. For best results, limit arguments to 50 characters.

An object's  **AddOnName** property indicates the name of the add-on to which the arguments are sent.

 Beginning with Visio 2002, the **AddOnName** property used in the following example cannot execute a string that contains arbitrary Microsoft Visual Basic code. To call code that in previous versions of Visio you would have passed to the **AddOnName** property, move it to a procedure in a document's Visual Basic project that is called from the **AddOnName** property, as shown in the following example.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to set the  **AddOnArgs** property of a menu item. It also shows how to add a menu and menu item to the drawing window menu set and how to set some of the menu item's other properties, such as **Caption** , **AddOnName** , and **ActionText** .

This example assumes that you already have a macro in the current Visual Basic project, and that that macro takes one argument called "Arg1". Before running this example, replace  _macroname_ with the name of your macro.

To restore the built-in user interface after running this macro, use the  **ThisDocument.ClearCustomMenus** method.




```vb
 
Public Sub AddOnArgs_Example() 
 
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


