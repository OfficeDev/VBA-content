---
title: Window.SelectedMasters Property (Visio)
keywords: vis_sdr.chm11651645
f1_keywords:
- vis_sdr.chm11651645
ms.prod: visio
api_name:
- Visio.Window.SelectedMasters
ms.assetid: 8a4546b4-4930-8c69-9df6-84e6b5a1bce0
ms.date: 06/08/2017
---


# Window.SelectedMasters Property (Visio)

 Returns an array of the masters or master shortcuts selected in a Microsoft Visio stencil window. Read-only.


## Syntax

 _expression_ . **SelectedMasters**

 _expression_ A variable that represents a **Window** object.


### Return Value

Object()


## Remarks

The  **SelectedMasters** property applies only to stencil windows. If you try to access the **SelectedMasters** property for other types of window, Visio might return an error.


## Example

This Microsoft Visual Basic for Applications (VBA) macro uses the  **SelectedMasters** property to get the number of masters and master shortcuts selected in a stencil window and then prints the name of the stencil and the selected masters and master shortcuts in the **Immediate** window.

Before running this macro, make sure that at least one master or master shortcut is selected in a docked stencil in the active Visio window.




```vb
Sub SelectedMasters_Example() 
 
 Dim vsoWindow As Visio.Window 
 Dim aobjSelectedMasters() As Object 
 Dim intNumberMasters As Integer 
 Dim intNumberMasterShortCuts As Integer 
 Dim vsoMaster As Visio.Master 
 Dim vsoMasterShortcut As Visio.MasterShortcut 
 intNumberMaster = 0 
 intNumberMasterShortCuts = 0 
 
 For Each vsoWindow In ActiveWindow.Windows 
 
 If (vsoWindow.Type = visDockedStencilBuiltIn) Then 
 aobjSelectedMasters = vsoWindow.SelectedMasters 
 
 For intCounter = LBound(aobjSelectedMasters) To UBound(aobjSelectedMasters) 
 On Error Resume Next 
 Set vsoMaster = Nothing 
 Set vsoMasterShortcut = Nothing 
 Set vsoMaster = aobjSelectedMasters(intCounter) 
 
 If Not vsoMaster Is Nothing Then 
 intNumberMasters = intNumberMasters + 1 
 Else 
 Set vsoMasterShortcut = aobjSelectedMasters(intCounter) 
 
 If Not vsoMasterShortcut Is Nothing Then 
 intNumberMasterShortCuts = intNumberMasterShortCuts + 1 
 End If 
 
 End If 
 
 Next 
 
 If (intNumberMasters > 0 Or intNumberMasterShortCuts > 0) Then 
 Debug.Print "The stencil " &; vsoWindow.Document.Name 
 Debug.Print "has" &; Str(intNumberMasters) &; " masters selected and " 
 Debug.Print Str(intNumberMasterShortCuts) &; " master shortcuts selected." 
 Exit For 
 End If 
 
 End If 
 
 Next 
 
End Sub
```


