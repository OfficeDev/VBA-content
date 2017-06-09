---
title: Windows.Add Method (Visio)
keywords: vis_sdr.chm11716685
f1_keywords:
- vis_sdr.chm11716685
ms.prod: visio
api_name:
- Visio.Windows.Add
ms.assetid: a4180d23-0333-046a-2c23-1a1f1b16240b
ms.date: 06/08/2017
---


# Windows.Add Method (Visio)

Adds a new  **Window** object to the **Windows** collection.


## Syntax

 _expression_ . **Add**( **_bstrCaption_** , **_nFlags_** , **_nType_** , **_nLeft_** , **_nTop_** , **_nWidth_** , **_nHeight_** , **_bstrMergeID_** , **_bstrMergeClass_** , **_nMergePosition_** )

 _expression_ A variable that represents a **Windows** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _bstrCaption_|Optional| **Variant**|The title of window; default is "Untitled".|
| _nFlags_|Optional| **Variant**| Initial window state. Can contain any combination of **[VisWindowStates](viswindowstates-enumeration-visio.md)** constants declared in the Visio type library; default varies based on the _nType_.|
| _nType_|Optional| **Variant**|Type of new window. Can be any one of the  **[VisWinTypes](viswintypes-enumeration-visio.md)** constants declared in the Visio type library. Defaults to **visStencilAddon** for **Application.Windows** ; defaults to **visAnchorBarAddon** for **Window.Windows**|
| _nLeft_|Optional| **Variant**|Position of the left side of the window.|
| _nTop_|Optional| **Variant**|Position of the top of the window.|
| _nWidth_|Optional| **Variant**|Width of the client area of the window.|
| _nHeight_|Optional| **Variant**|Height of the client area of the window.|
| _bstrMergeID_|Optional| **Variant**|Merge ID of the window.|
| _bstrMergeClass_|Optional| **Variant**|Merge class of the window.|
| _nMergePosition_|Optional| **Variant**|Merge position of the window.|

### Return Value

Window


## Remarks

Use this method to get an empty parent frame window within the Visio window space that you can populate with child windows. You must be in the Visio process space (for example, in a DLL or VSL-based add-on) to use the  **Window** object returned by this method as a parent to your windows.

Use the value returned by the  **WindowHandle32** property as an **HWND** for use as a parent to your own windows.


## Example

The following macro shows how to use the  **Add** method to add a **Window** object to the **Windows** collection. It creates a new, empty parent frame window, docked to the bottom of the drawing window. Then it populates the new parent frame window with a child window, in this case a form, so that the new window does not appear empty.

Add a form to your Microsoft Visual Basic (VBA) project called  **frmMain**, and then add a  **TextBox** control named **txtForm** to the form.

The  **SetParent** , **FindWindow** , and **SetWindowLongLib** functions are from the Windows API, and are necessary to add the form to the new window.

Add the following code to the form module to resize the text box when the form is resized:




```vb
Private Sub UserForm_Resize() 
 txtForm.Width = txtForm.Parent.Width - 10 
 txtForm.Height = txtForm.Parent.Height - 10 
End Sub
```

Then add the following code to the document project:




```vb
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long 
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long 
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long 
 
Private Const GWL_STYLE = (-16) 
Private Const WS_CHILD = &;H40000000 
Private Const WS_VISIBLE = &;H10000000 
 
Public Sub AddWindow_Example() 
 
 Dim vsoWindow As Visio.Window 
 Dim frmNewWindow As UserForm 
 Dim lngFormHandle As Long 
 
 'Add a new Anchor Bar window docked to the bottom of the Visio drawing window 
 Set vsoWindow = ActiveWindow.Windows.Add("My New Window", visWSVisible + visWSDockedBottom, visAnchorBarAddon, , , 300, 210) 
 
 'Create a new windows form 
 Set frmNewWindow = New frmMain 
 
 'Get the 32-bit handle of the new window. 
 lngFormHandle = FindWindow(vbNullString, "My New Window") 
 
 SetWindowLong lngFormHandle, GWL_STYLE, WS_CHILD Or WS_VISIBLE 
 SetParent lngFormHandle, vsoWindow.WindowHandle32 
 
End Sub
```


