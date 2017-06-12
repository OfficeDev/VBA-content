---
title: ToolbarItem.CntrlType Property (Visio)
keywords: vis_sdr.chm13513265
f1_keywords:
- vis_sdr.chm13513265
ms.prod: visio
api_name:
- Visio.ToolbarItem.CntrlType
ms.assetid: 88995561-9227-61ae-693a-83f5bba5bede
ms.date: 06/08/2017
---


# ToolbarItem.CntrlType Property (Visio)

Gets or sets the control type of a menu, menu item, or toolbar item. Read/write.


## Syntax

 _expression_ . **CntrlType**

 _expression_ A variable that represents a **ToolbarItem** object.


### Return Value

Integer


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

If you are adding a custom toolbar button, set the  **CntrlType** property to **visCtrlTypeBUTTON** . The following table describes the control types declared by the Visio type library in **VisUICtrlTypes** .



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visCtrlTypeSPLITBUTTON_MRU_COMMAND**|18|Split button, with MRU command behavior.|
| **visCtrlTypeBUTTON_OWNERDRAW**|33|Owner-draw push button.|
| **visCtrlTypeDROPDOWN**|272|Drop-down combo box.|
| **visCtrlTypeEDITBOX**|64|Text box.|
| **visCtrlTypeCOMBOBOX**|128|Combo box.|
| **visCtrlTypeBUTTON**|2|Push button.|
| **visCtrlTypeSPLITBUTTON_MRU_COLOR**|16|Split button, with MRU color behavior.|
| **visCtrlTypeLABEL**|2048|Label.|
| **visCtrlTypeSPLITBUTTON**|17|Split button.|

## Example

The following two examples demonstrate different ways to use the  **CntrlType** property in your programs.

This first example shows how to use the  **CntrlType** property to set the type of a new toolbar item. The example adds a custom toolbar to the **Toolbars** collection, and then adds a button to the toolbar. The toolbar appears in the Visio user interface and is available while the document is active.




```vb
 
Sub CntrlType_Example1() 
 
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
 Set vsoToolbars = vsoUIObject.ToolbarSets.ItemAtID( _ 
 Visio.visUIObjSetDrawing).Toolbars 
 
 'Add a toolbar to the collection. 
 Set vsoToolbar = vsoToolbars.Add 
 
 'Add an item to the toolbar. 
 Set vsoToolbarItem = vsoToolbar.ToolbarItems.Add 
 
 With vsoToolbarItem 
 
 'Set the new item to be a button. 
 .CntrlType = Visio.visCtrlTypeBUTTON 
 
 'Set the icon of the new button. 
 .FaceID = Visio.visIconIXCUSTOM_CARDS 
 
 'Set the CmdNum property of the new button 
 .CmdNum = 1 
 
 'Set the Width property of the new button 
 'wide enough that the toolbar name is readable. 
 .Width = 100 
 
 End With 
 
 'Tell Microsoft Office Visio to use the new UI object while 
 'this document is active. 
 ThisDocument.SetCustomToolbars vsoUIObject 
 
End Sub
```

This second example shows another way to use the  **CntrlType** property to add a toolbar button and set the button icon.

Before running this macro, replace  _fullpath\filename_ in the code below with the full path to and file name of an icon file (.ico) on your computer. To restore the built-in toolbars in Visio after you run this macro, call the **ThisDocument.ClearCustomToolbars** method.




```vb
 
Public Sub CntrlType_Example2() 
 
 Dim vsoUIObject As Visio.UIObject 
 Dim vsoToolbarSet As Visio.ToolbarSet 
 Dim vsoToolbarItems As Visio.ToolbarItems 
 Dim vsoToolbarItem As Visio.ToolbarItem 
 
 'Get the UIObject object for the copy of the Microsoft Office toolbars. 
 Set vsoUIObject = Visio.Application.BuiltInToolbars(0) 
 
 'Get the drawing window toolbar sets. 
 'NOTE: Use ItemAtID to get the toolbar set. 
 'Using vsoUIObject.ToolbarSets(visUIObjSetDrawing) will not work. 
 Set vsoToolbarSet = vsoUIObject.ToolbarSets.ItemAtID(visUIObjSetDrawing) 
 
 'Get the ToolbarItems collection of the first toolbar in the collection. 
 Set vsoToolbarItems = vsoToolbarSet.Toolbars(0).ToolbarItems 
 
 'Add a new toolbar item in the first position on the toolbar. 
 Set vsoToolbarItem = vsoToolbarItems.AddAt(0) 
 
 'Make the new toolbar item a button. 
 'This requires setting both the CmdNum and CntrlType properties 
 vsoToolbarItem.CmdNum = 1 
 vsoToolbarItem.CntrlType = visCtrlTypeBUTTON 
 
 'Set the toolbar button icon. 
 vsoToolbarItem.IconFileName "fullpath\filename " 
 
 'Tell Visio to actually use the new custom UI. 
 ThisDocument.SetCustomToolbars vsoUIObject 
 
End Sub
```


