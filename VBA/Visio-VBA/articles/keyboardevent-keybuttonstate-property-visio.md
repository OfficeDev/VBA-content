---
title: KeyboardEvent.KeyButtonState Property (Visio)
keywords: vis_sdr.chm17051715
f1_keywords:
- vis_sdr.chm17051715
ms.prod: visio
api_name:
- Visio.KeyboardEvent.KeyButtonState
ms.assetid: c2ab3fa3-39c6-fb34-1f56-342cf080d9d5
ms.date: 06/08/2017
---


# KeyboardEvent.KeyButtonState Property (Visio)

Returns the state of mouse buttons and the SHIFT and CTRL keys associated with a keyboard event. Read-only.


## Syntax

 _expression_ . **KeyButtonState**

 _expression_ A variable that represents a **KeyboardEvent** object.


### Return Value

Long


## Remarks

Possible values for  **KeyButtonState** can be a combination of any of the values shown in the following table, which are declared in **VisKeyButtonFlags** in the Visio type library. For example, if **KeyButtonState** returns 9, it indicates that the user clicked the left mouse button while pressing CTRL.



|**Constant **|**Value **|
|:-----|:-----|
| **visKeyControl**|8|
| **visKeyShift**|4|
| **visMouseLeft**|1|
| **visMouseMiddle**|16|
| **visMouseRight**|2|

## Example

This class module shows how to define a sink class called  **KeyboardListener** that listens for events fired by keyboard actions in the active window. It declares the object variable _vsoWindow_ by using the **WithEvents** keyword. The class module also contains event handlers for the **KeyDown** , **KeyPress** , and **KeyUp** events.

To run this example, insert a new class module in your Microsoft Visual Basic for Applications (VBA) project, name it  **KeyboardListener** , and insert the following code in the module.




```vb
Dim WithEvents vsoWindow As Visio.Window 
 
Private Sub Class_Initialize() 
 
 Set vsoWindow = ActiveWindow 
 
End Sub 
 
Private Sub Class_Terminate() 
 
 Set vsoWindow = Nothing 
 
End Sub 
 
Private Sub vsoWindow_KeyDown(ByVal KeyCode As Long, ByVal KeyButtonState As Long, CancelDefault As Boolean) 
 
 Debug.Print "KeyCode is "; KeyCode 
 Debug.Print "KeyButtonState is" ; KeyButtonState 
 
End Sub 
 
Private Sub vsoWindow_KeyPress(ByVal KeyAscii As Long, CancelDefault As Boolean) 
 
 Debug.Print "KeyAscii value is "; KeyAscii 
 
End Sub 
 
Private Sub vsoWindow_KeyUp(ByVal KeyCode As Long, ByVal KeyButtonState As Long, CancelDefault As Boolean) 
 
 Debug.Print "KeyCode is "; KeyCode 
 Debug.Print "KeyButtonState is" ; KeyButtonState 
 
End Sub
```

Then, insert the following code in the  **ThisDocument** project.




```vb
Dim myKeyboardListener As KeyboardListener 
 
Private Sub Document_DocumentSaved(ByVal doc As IVDocument) 
 
 Set myKeyboardListener = New KeyboardListener 
 
End Sub 
 
Private Sub Document_BeforeDocumentClose(ByVal doc As IVDocument) 
 
 Set myKeyboardListener = Nothing 
 
End Sub
```




```

```

Save the document to initialize the class, and then press any key to fire a  **KeyDown** event. In the Immediate window, the handler prints the code of the key that was pressed to fire the event and the state of the SHIFT and CTRL keys at the time the event fired.


