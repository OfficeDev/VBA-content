---
title: ToolbarItem.PaletteWidth Property (Visio)
keywords: vis_sdr.chm13514010
f1_keywords:
- vis_sdr.chm13514010
ms.prod: visio
api_name:
- Visio.ToolbarItem.PaletteWidth
ms.assetid: bb3356d9-bfa3-c2e5-129f-70c0b225add6
ms.date: 06/08/2017
---


# ToolbarItem.PaletteWidth Property (Visio)

Gets or sets the width of a palette in pixels. Read/write.


## Syntax

 _expression_ . **PaletteWidth**

 _expression_ A variable that represents a **ToolbarItem** object.


### Return Value

Integer


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

A palette, like a toolbar, is organized horizontally, and items wrap to a new row if there is not enough horizontal space available. By default, only the icons of the items are shown.


## Example

This example shows how to use the  **PaletteWidth** property to set the width of a palette on a custom toolbar. The example adds a custom toolbar to the **Toolbars** collection and then adds a button to the toolbar. The toolbar appears in the Microsoft Visio user interface and is available while the document is active.

To restore the built-in toolbars in Visio after you run this macro, call the  **ThisDocument.ClearCustomToolbars** method.




```vb
 
Sub PaletteWidth_Example1() 
 
 Dim vsoUIObject As Visio.UIObject 
 Dim vsoToolbars As Visio.Toolbars 
 Dim vsoToolbar As Visio.Toolbar 
 Dim vsoToolbarItems As Visio.ToolbarItems 
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
 Set vsoToolbars = vsoUIObject.ToolbarSets.ItemAtID( _ 
 Visio.visUIObjSetDrawing).Toolbars 
 
 'Add a toolbar to the collection. 
 Set vsoToolbar = vsoToolbars.Add 
 With vsoToolbar 
 
 'Set the title of the toolbar. 
 .Caption = "My New Toolbar" 
 
 'Float the toolbar at coordinates (300,200). 
 .Position = Visio.visBarFloating 
 .Left = 300 
 .Top = 200 
 
 'Disallow docking the new toolbar. 
 .Protection = Visio.visBarNoHorizontalDock _ 
 + Visio.visBarNoVerticalDock 
 
 End With 
 
 'Add an item to the toolbar. 
 Set vsoToolbarItem = vsoToolbar.ToolbarItems.Add 
 With vsoToolbarItem 
 
 'Set the new item to be a button. 
 'Because the CmdNum property isn't specified, 
 'the button type defaults to palette. 
 .CntrlType = Visio.visCtrlTypeBUTTON 
 
 'Set the icon of the new button. 
 .FaceID = Visio.visIconIXCUSTOM_CARDS 
 
 'Make the button wide enough that 
 'we can read the toolbar name. 
 .Width = 150 
 
 'Set the PaletteWidth property of the new button. 
 .PaletteWidth = 50 
 
 End With 
 
 'Add some buttons to the palette. 
 Set vsoToolbarItems = vsoToolbarItem.ToolbarItems 
 With vsoToolbarItems.Add 
 .FaceID = Visio.visIconIXCUSTOM_SPADE 
 .CmdNum = Visio.visCmdFileSave 
 End With 
 With vsoToolbarItems.Add 
 .FaceID = Visio.visIconIXCUSTOM_DIAMOND 
 .CmdNum = Visio.visCmdFilePrint 
 End With 
 With vsoToolbarItems.Add 
 .FaceID = Visio.visIconIXCUSTOM_CLUB 
 .CmdNum = Visio.visCmdToolsRunVBE 
 End With 
 With vsoToolbarItems.Add 
 .FaceID = Visio.visIconIXCUSTOM_HEART 
 .CmdNum = Visio.visCmdToolsMacroDlg 
 End With 
 
 'Use the new UIObject object while 
 'this document is active. 
 ThisDocument.SetCustomToolbars vsoUIObject 
 
End Sub
```


