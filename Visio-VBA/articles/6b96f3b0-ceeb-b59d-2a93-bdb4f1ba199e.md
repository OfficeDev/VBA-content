
# KeyboardEvent.KeyCode Property (Visio)

 **Last modified:** July 28, 2015

 _**Applies to:** Visio 2013 Preview_

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns the key associated with  **KeyDown** and **KeyUp** events. Read-only.

## Syntax
<a name="sectionSection0"> </a>

 _expression_. **KeyCode**

 _expression_A variable that represents a  **KeyboardEvent** object.


### Return Value

Long


## Remarks
<a name="sectionSection1"> </a>

Possible values for  **KeyCode** are declared in **KeyCodeConstants** in the Microsoft Visual Basic for Applications (VBA) library.


## Example
<a name="sectionSection2"> </a>

This class module shows how to define a sink class called  **KeyboardListener** that listens for events fired by keyboard actions in the active window. It declares the object variable _vsoWindow_ by using the **WithEvents** keyword. The class module also contains event handlers for the **KeyDown**,  **KeyPress**, and  **KeyUp** events.

To run this example, insert a new class module in your VBA project, name it  **KeyboardListener**, and insert the following code in the module.




```
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




```
Dim myKeyboardListener As KeyboardListener 
 
Private Sub Document_DocumentSaved(ByVal doc As IVDocument) 
 
 Set myKeyboardListener = New KeyboardListener 
 
End Sub 
 
Private Sub Document_BeforeDocumentClose(ByVal doc As IVDocument) 
 
 Set myKeyboardListener = Nothing 
 
End Sub 
 

```

Save the document to initialize the class, and then press any key to fire a  **KeyDown** event. In the Immediate window, the handler prints the code of the key that was pressed to fire the event and the state of the SHIFT and CTRL keys at the time the event fired.

