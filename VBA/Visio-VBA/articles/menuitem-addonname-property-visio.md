---
title: MenuItem.AddOnName Property (Visio)
keywords: vis_sdr.chm12913050
f1_keywords:
- vis_sdr.chm12913050
ms.prod: visio
api_name:
- Visio.MenuItem.AddOnName
ms.assetid: dfe65141-f5e4-77b3-8113-4650a602ea34
ms.date: 06/08/2017
---


# MenuItem.AddOnName Property (Visio)

Gets or sets the name of an add-on or procedure that is run when its associated menu item is selected. Read/write.


## Syntax

 _expression_ . **AddOnName**

 _expression_ A variable that represents a **MenuItem** object.


### Return Value

String


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

Assuming that the name of the add-on in the  **Addons** collection is _string_ , if the project of the currently active document (or another project if it is referenced) does not have a procedure named _string_ , or if the arguments passed in _string_ do not match those specified in the procedure, Microsoft Visio runs the add-on named _string_ . If no add-on named _string_ can be found, Visio does nothing and reports no error. (You can use the **TraceFlags** property to monitor the procedures and add-ons that Visio attempts to run.)

If  _string_ is an add-on, use the **AddOnArgs** property to specify arguments to send to the add-on when it is run.

If  _string_ is a procedure, specify arguments using _procname(arguments)_ or _procname arguments_ .

When calling a procedure in a standard module it is recommended that you prefix the string with the module name that contains the procedure (for example,  _moduleName.procName_ ) because more than one module can have a procedure with the same name.

To call a procedure in a project other than the project of the active document, use the syntax  _projName.modName.procName_ (you must have explicitly set a reference to _projName_ in your Visual Basic project).

If the  **AddOnName** property is set, Visio ignores the object's **CmdNum** property.


 **Note**  Beginning with Visio 2002, the  **AddOnName** property cannot execute a string that contains arbitrary VBA code. To call code that in previous versions of Visio you would have passed to the **AddOnName** property, move the code to a procedure in a document's VBA project that is called from the **AddOnName** property.


## Example

This VBA macro shows how to set the  **AddOnName** property of a menu item. It also shows how to add a menu and menu item to the **Add-ins** tab, and how to set some of the menu item's other properties, such as **Caption** , **AddOnArgs** , and **ActionText** .

This example assumes that you already have a macro named  _macroname_ in the project of the active document, and that the macro takes an argument called "Arg1." Before running this example, replace _macroname_ with the name of your macro.

To restore the built-in user interface in Microsoft Visio after you run this macro, call the  **ThisDocument.ClearCustomMenus** method.




```vb
 
Public Sub AddOnName_Example() 
 
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


