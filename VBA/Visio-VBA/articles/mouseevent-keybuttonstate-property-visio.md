---
title: MouseEvent.KeyButtonState Property (Visio)
keywords: vis_sdr.chm17151715
f1_keywords:
- vis_sdr.chm17151715
ms.prod: visio
api_name:
- Visio.MouseEvent.KeyButtonState
ms.assetid: d4a408af-38a4-6e3f-3dfc-6ebf342c6bb1
ms.date: 06/08/2017
---


# MouseEvent.KeyButtonState Property (Visio)

Returns the state of mouse buttons and the SHIFT and CTRL keys associated with a mouse event. Read-only.


## Syntax

 _expression_ . **KeyButtonState**

 _expression_ A variable that represents a **MouseEvent** object.


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

This class module shows how to define a sink class called  **MouseListener** that listens for events fired by mouse actions in the active window. It declares the object variable _vsoWindow_ by using the **WithEvents** keyword. The class module also contains an event handler for the **MouseDown** event that prints to the Immediate window the state of the mouse buttons and CTRL and SHIFT keys when the event fired.

To run this example, insert a new class module in your VBA project, name it  **MouseListener** , and insert the following code in the module.




```vb
Dim WithEvents vsoWindow As Visio.Window 
 
Private Sub Class_Initialize() 
 
 Set vsoWindow = ActiveWindow 
 
End Sub 
 
Private Sub Class_Terminate() 
 
 Set vsoWindow = Nothing 
 
End Sub 
 
Private Sub vsoWindow_MouseDown(ByVal Button As Long, ByVal KeyButtonState As Long, ByVal x As Double, ByVal y As Double, CancelDefault As Boolean) 
 
 Debug.Print "KeyButtonState is"; KeyButtonState 
 
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

Save the document to initialize the class, and then click anywhere in the active window (optionally, while pressing SHIFT and/or CTRL) to fire a  **MouseDown** event. In the Immediate window, the handler prints the name of the mouse button that was clicked to fire the event. If you pressed either or both of the keys, the name of the key or keys you pressed will print as well.


