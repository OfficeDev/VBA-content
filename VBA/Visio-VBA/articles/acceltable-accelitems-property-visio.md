---
title: AccelTable.AccelItems Property (Visio)
keywords: vis_sdr.chm14751180
f1_keywords:
- vis_sdr.chm14751180
ms.prod: visio
api_name:
- Visio.AccelTable.AccelItems
ms.assetid: 700cee8b-7521-8214-b83b-731dd91429ac
ms.date: 06/08/2017
---


# AccelTable.AccelItems Property (Visio)

Returns the  **AccelItems** collection of an **AccelTable** object. Read-only.


## Syntax

 _expression_ . **AccelItems**

 _expression_ A variable that represents an **AccelTable** object.


### Return Value

AccelItems


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **AccelItems** property to delete an accelerator key from a built-in menu.

To restore the built-in menus in Microsoft Visio after you run this macro, call the  **ThisDocument.ClearCustomMenus** method.




```vb
 
Public Sub AccelItems_Example() 
 
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


