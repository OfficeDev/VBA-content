---
title: ToolbarItem.FaceID Property (Visio)
keywords: vis_sdr.chm13513495
f1_keywords:
- vis_sdr.chm13513495
ms.prod: visio
api_name:
- Visio.ToolbarItem.FaceID
ms.assetid: 226b00fd-c8ab-3a6a-811f-ebf8e1364526
ms.date: 06/08/2017
---


# ToolbarItem.FaceID Property (Visio)

Gets or sets the icon for an item. Read/write.


## Syntax

 _expression_ . **FaceID**

 _expression_ A variable that represents a **ToolbarItem** object.


### Return Value

Integer


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

You can use any of the constants prefixed with  **visIconIX** that are declared by the Visio type library in **[VisUIIconIDs](visuiiconids-enumeration-visio.md)** .

The  **FaceID** property determines a button's icon, but not its function. Use the **CmdNum** property of a **ToolbarItem** object to set a button's function.

The  **FaceID** property is the same as the **TypeSpecific1** property when the **CtrlType** property is type **visCtrlTypeBUTTON** , which is declared in the Visio type library in **[VisUICtrlTypes](visuictrltypes-enumeration-visio.md)** .


## Example

This example adds a custom toolbar to the  **Toolbars** collection and then adds a button to the toolbar. The example shows how to use the **FaceID** property to assign the icon for the button. This toolbar appears in the Microsoft Visio user interface and is available while the document is active.

To restore the built-in toolbars in Microsoft Visio after you run this macro, call the  **ThisDocument.ClearCustomToolbars** method.




```vb
 
Sub FaceID_Example() 
 
 Dim vsoUIObject As Visio.UIObject 
 Dim vsoToolbars As Visio.Toolbars 
 Dim vsoToolbar As Visio.Toolbar 
 Dim vsoToolbarItem As Visio.ToolbarItem 
 
 'Check whether there are document custom toolbars. 
 If ThisDocument.CustomToolbars Is Nothing Then 
 
 'Check whether there are application custom toolbars. 
 If Visio.Application.CustomToolbars Is Nothing Then 
 
 'Use the built-in toolbars. 
 Set vsoUIObject = Visio.Application.BuiltInToolbars(0) 
 
 Else 
 
 'Use the application custom toolbars. 
 Set vsoUIObject = Visio.Application.CustomToolbars.Clone 
 
 End If 
 
 Else 
 
 'Use the document custom toolbars. 
 Set vsoUIObject = ThisDocument.CustomToolbars 
 
 End If 
 
 'Get the Toolbars collection for the drawing window context. 
 Set VsoToolbars = vsoUIObject.ToolbarSets.ItemAtID( _ 
 Visio.visUIObjSetDrawing).Toolbars 
 
 'Add a toolbar to the collection. 
 Set vsoToolbar = vsoToolbars.Add 
 
 'Set the title of the toolbar. 
 vsoToolbar.Caption = "Example" 
 
 'Add an item to the toolbar. 
 Set vsoToolbarItem = vsoToolbar.ToolbarItems.Add 
 With vsoToolbarItem 
 
 'Set the new item to be a button. 
 .CntrlType = Visio.visCtrlTypeBUTTON 
 
 'Set the icon of the new button. 
 .FaceID = Visio.visIconIXCUSTOM_CARDS 
 
 'Set the CmdNum property of the new button. 
 .CmdNum = 1 
 
 'Set the Width property of the new button 
 'wide enough that the toolbar name is readable. 
 .Width = 100 
 
 End With 
 
 'Tell Visio to use the new UIObject object while 
 'this document is active. 
 ThisDocument.SetCustomToolbars vsoUIObject 
 
End Sub
```


