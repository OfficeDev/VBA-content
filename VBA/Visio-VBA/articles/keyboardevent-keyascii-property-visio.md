---
title: KeyboardEvent.KeyAscii Property (Visio)
keywords: vis_sdr.chm17051720
f1_keywords:
- vis_sdr.chm17051720
ms.prod: visio
api_name:
- Visio.KeyboardEvent.KeyAscii
ms.assetid: 0e4e1b3b-a93a-20f3-982f-88879e2a6393
ms.date: 06/08/2017
---


# KeyboardEvent.KeyAscii Property (Visio)

Returns the ASCII code associated with a  **KeyPress** event. Read-only.


## Syntax

 _expression_ . **KeyAscii**

 _expression_ A variable that represents a **KeyboardEvent** object.


### Return Value

Long


## Remarks

The values returned by  **KeyAscii** are ASCII codes. To see a list of these codes, search for "ASCII character codes" on MSDN, the Microsoft Developer Network.


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

Save the document to initialize the class, and then press any key to fire a  **KeyPress** event. In the Immediate window, the handler prints the ASCII code of the key that was pressed to fire the event.


