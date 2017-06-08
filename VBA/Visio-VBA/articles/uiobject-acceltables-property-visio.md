---
title: UIObject.AccelTables Property (Visio)
keywords: vis_sdr.chm14913005
f1_keywords:
- vis_sdr.chm14913005
ms.prod: visio
api_name:
- Visio.UIObject.AccelTables
ms.assetid: 01cdfc77-47b3-b160-fbaa-9e7d615abff2
ms.date: 06/08/2017
---


# UIObject.AccelTables Property (Visio)

Returns the  **AccelTables** collection of a **UIObject** object. Read-only.


## Syntax

 _expression_ . **AccelTables**

 _expression_ A variable that represents a **UIObject** object.


### Return Value

AccelTables


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

If a  **UIObject** object represents menu items and accelerators (for example, if you used the **BuiltInMenus** property of an **Application** object to retrieve the **UIObject** object), its **AccelTables** collection represents tables of accelerator keys for that **UIObject** object.

To retrieve accelerators for a particular window context, for example, the drawing window, use the  **ItemAtID** property of an **AccelTables** collection. If a window context does not include accelerators, it has no **AccelTables** collection. Valid window context IDs are declared in **[VisUIObjSets](visuiobjsets-enumeration-visio.md)** in the Visio type library.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **AccelTables** property to delete an accelerator key from a built-in menu.

To restore the built-in menus in Microsoft Visio after you run this macro, call the  **ThisDocument.ClearCustomMenus** method.




```vb
 
Public Sub AccelTables_Example() 
 
 Dim vsoUIObject As Visio.UIObject 
 Dim vsoAccelTable As Visio.AccelTable 
 Dim vsoAccelItems As Visio.AccelItems 
 Dim vsoAccelItem As Visio.AccelItem 
 Dim intCounter As Integer 
 
 'Retrieve the UIObject object for the copy of the built-in menus. 
 Set vsoUIObject = Visio.Application.BuiltInMenus 
 
 'Set vsoAccelTable to the drawing menu set. 
 Set vsoAccelTable = vsoUIObject.AccelTables.ItemAtID(visUIObjSetDrawing) 
 
 'Retrieve the accelerator items collection. 
 Set vsoAccelItems = vsoAccelTable.AccelItems 
 
 'Retrieve the accelerator item for the Visual Basic Editor. 
 'To do this, we must iterate through the collection 
 'and locate the item we want to manipulate. 
 'The item can be identified either by checking 
 'the CmdNum property or by checking for the specific key. 
 'Because checking for the key requires looking at the Alt, 
 'Control, Shift, and Key properties, it is better to use the 
 'CmdNum property. Because we retrieved the built-in menus, 
 'we know that we can find the accelerator. 
 For intCounter = 0 To vsoAccelItems.Count - 1 
 Set vsoAccelItem = vsoAccelItems.Item(intCounter) 
 If vsoAccelItem.CmdNum = Visio.visCmdToolsRunVBE Then 
 Exit For 
 
 End If 
 Next intCounter 
 
 'Delete the accelerator. 
 vsoAccelItem.Delete 
 
 'Tell Visio to use the new UI. 
 ThisDocument.SetCustomMenus vsoUIObject 
 
End Sub
```


