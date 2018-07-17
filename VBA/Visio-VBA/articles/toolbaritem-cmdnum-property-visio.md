---
title: ToolbarItem.CmdNum Property (Visio)
keywords: vis_sdr.chm13513255
f1_keywords:
- vis_sdr.chm13513255
ms.prod: visio
api_name:
- Visio.ToolbarItem.CmdNum
ms.assetid: 69be3d63-a149-60ff-081e-fa20d8650685
ms.date: 06/08/2017
---


# ToolbarItem.CmdNum Property (Visio)

Gets or sets the command ID associated with a toolbar item. Read/write.


## Syntax

 _expression_ . **CmdNum**

 _expression_ A variable that represents a **ToolbarItem** object.


### Return Value

Integer


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

When the  **AddOnName** property of a **ToolbarItem** object indicates an add-on to run, Microsoft Visio automatically assigns a **CmdNum** property.

To insert a separator in a spacer in a toolbar preceding a  **ToolbarItem** object, use the **BeginGroup** property.

Valid command IDs are declared by the Visio type library in  **[VisUICmds](visuicmds-enumeration-visio.md)** . They have the prefix **visCmd** .


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how use the  **CmdNum** property to get a particular built-in Visio toolbar button, and then it shows how to change the button's icon. The new icon persists as long as the document is active.

This macro assumes you are not using a custom user interface (UI).

Before running this macro, replace  _fullpath\filename_ in the following code with the full path to and file name of an icon file (.ico) on your computer.

To restore the built-in toolbars in Visio after you run this macro, call the  **ThisDocument.ClearCustomToolbars** method.




```vb
 
Public Sub CmdNum_Example() 
 
 Dim vsoUIObject As Visio.UIObject 
 Dim vsoToolbarSet As Visio.ToolbarSet 
 Dim vsoToolbarItems As Visio.ToolbarItems 
 Dim vsoToolbarItem As Visio.ToolbarItem 
 Dim intCounter As Integer 
 Dim blsFound As Boolean 
 
 'Get the UIObject object for the copy of the Microsoft Office toolbars. 
 Set vsoUIObject = Visio.Application.BuiltInToolbars(0) 
 
 'Get the drawing window toolbar sets. 
 'NOTE: Use ItemAtID to get the toolbar set. 
 'Using vsoUIObject.ToolbarSets(visUIObjSetDrawing) will not work. 
 Set vsoToolbarSet = vsoUIObject.ToolbarSets.ItemAtID(visUIObjSetDrawing) 
 
 'Get the vsoToolbarItems collection. 
 Set vsoToolbarItems = vsoToolbarSet.Toolbars(0).ToolbarItems 
 
 'Get the toolbar item for the Save toolbar button. 
 blsFound = False 
 For intCounter = 0 To vsoToolbarItems.Count - 1 
 
 Set vsoToolbarItem = vsoToolbarItems(intCounter) 
 If vsoToolbarItem.CmdNum = visCmdFileSave Then 
 blsFound = True 
 Exit For 
 
 End If 
 
 Next intCounter 
 
 If blsFound Then 
 
 'Set the icon. 
 vsoToolbarItem.IconFileName "fullpath\filename"  
 'Indicate to Visio to use the new custom UI. 
 ThisDocument.SetCustomToolbars vsoUIObject 
 
 End If 
 
End Sub
```


