---
title: MouseEvent.Button Property (Visio)
keywords: vis_sdr.chm17151905
f1_keywords:
- vis_sdr.chm17151905
ms.prod: visio
api_name:
- Visio.MouseEvent.Button
ms.assetid: e89c51a3-52f5-348c-e3de-2b2459701bfb
ms.date: 06/08/2017
---


# MouseEvent.Button Property (Visio)

Returns the mouse button that was clicked to fire a  **MouseDown** or **MouseUp** event. Read-only.


## Syntax

 _expression_ . **Button**

 _expression_ An expression that returns a **MouseEvent** object.


### Return Value

Long


## Remarks

Possible values for the  **Button** property can be any of the constants shown in the following table, which are declared in **VisKeyButtonFlags** in the Visio type library.



|**Constant**|**Value**|
|:-----|:-----|
| **visMouseLeft**|1|
| **visMouseMiddle**|16|
| **visMouseRight**|2|

## Example

This class module shows how to define a sink class called  **MouseListener** that listens for events fired by mouse actions in the active window. It declares the object variable _vsoWindow_ by using the **WithEvents** keyword. The class module also contains event handlers for the **MouseDown** , **MouseMove** , and **MouseUp** events.

To run this example, insert a new class module in your Microsoft Visual Basic for Applications (VBA) project, name it  **MouseListener** , and insert the following code in the module.




```vb
Dim WithEvents vsoWindow As Window 
 
Private Sub Class_Initialize() 
 
 Set vsoWindow = ActiveWindow 
 
End Sub 
 
Private Sub Class_Terminate() 
 
 Set vsoWindow = Nothing 
 
End Sub 
 
Private Sub vsoWindow_MouseDown(ByVal Button As Long, ByVal KeyButtonState As Long, ByVal x As Double, ByVal y As Double, CancelDefault As Boolean) 
 
 Debug.Print "Button is: "; Button 
 
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

Save the document to initialize the class, and then click anywhere in the active window to fire a  **MouseDown** event. In the Immediate window, the handler prints the value that represents the mouse button that was clicked to fire the event.


