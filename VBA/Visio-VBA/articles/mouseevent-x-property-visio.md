---
title: MouseEvent.x Property (Visio)
keywords: vis_sdr.chm17151475
f1_keywords:
- vis_sdr.chm17151475
ms.prod: visio
api_name:
- Visio.MouseEvent.x
ms.assetid: baf35c3b-8548-68e0-733c-5a8385c42ebc
ms.date: 06/08/2017
---


# MouseEvent.x Property (Visio)

Returns the x-coordinate of the location in the Microsoft Visio window where a  **MouseDown** , **MouseMove** , or **MouseUp** event fired. Read-only.


## Syntax

 _expression_ . **x**

 _expression_ A variable that represents a **MouseEvent** object.


### Return Value

VisStatCodes


## Remarks

The  **x** property returns a value in internal drawing units.


## Example

This class module shows how to define a sink class called  **MouseListener** that listens for events fired by mouse actions in the active window. It declares the object variable _vsoWindow_ by using the **WithEvents** keyword. The class module also contains event handlers for the **MouseDown** , **MouseMove** , and **MouseUp** events.

To run this example, insert a new class module in your Microsoft Visual Basic for Applications (VBA) project, name it  **MouseListener** , and insert the following code in the module.




```vb
Dim WithEvents vsoWindow As Visio.Window 
 
Private Sub Class_Initialize() 
 
 Set vsoWindow = ActiveWindow 
 
End Sub 
 
Private Sub Class_Terminate() 
 
 Set vsoWindow = Nothing 
 
End Sub 
 
Private Sub vsoWindow_MouseDown(ByVal Button As Long, ByVal KeyButtonState As Long, ByVal x As Double, ByVal y As Double, CancelDefault As Boolean) 
 
 Debug.Print "x is: "; x 
 Debug.Print "y is: "; y 
 
End Sub 
 
Private Sub vsoWindow_MouseMove(ByVal Button As Long, ByVal KeyButtonState As Long, ByVal x As Double, ByVal y As Double, CancelDefault As Boolean) 
 
 Debug.Print "x-position is "; x 
 Debug.Print "y-position is "; y 
 
End Sub 
 
Private Sub vsoWindow_MouseUp(ByVal Button As Long, ByVal KeyButtonState As Long, ByVal x As Double, ByVal y As Double, CancelDefault As Boolean) 
 
 If Button = 1 Then 
 
 Debug.Print "Left mouse button released" 
 
 ElseIf Button = 2 Then 
 
 Debug.Print "Right mouse button released" 
 
 ElseIf Button = 16 Then 
 
 Debug.Print "Center mouse button released" 
 
 End If 
 
End Sub
```

Then, insert the following code in the  **ThisDocument** project.




```vb
Dim myMouseListener As MouseListener 
 
Private Sub Document_DocumentSaved(ByVal doc As IVDocument) 
 
 Set myMouseListener = New MouseListener 
 
End Sub 
 
Private Sub Document_BeforeDocumentClose(ByVal doc As IVDocument) 
 
 Set myMouseListener = Nothing 
 
End Sub
```

Save the document to initialize the class, and then click anywhere in the active window to fire a  **MouseDown** event. In the Immediate window, the handler prints the _x_ and _y_ coordinates of the location in the Visio window coordinate space where the mouse was clicked.


