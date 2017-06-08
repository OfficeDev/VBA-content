---
title: InvisibleApp.MouseMove Event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.MouseMove
ms.assetid: 7f360b69-539e-bdf2-2ff5-1b67090e01ac
ms.date: 06/08/2017
---


# InvisibleApp.MouseMove Event (Visio)

Occurs when the mouse is moved.


## Syntax

Private Sub  _expression_ _**MouseMove**( **_ByVal Button As Long_** , **_ByVal KeyButtonState As Long_** , **_ByVal x As Double_** , **_ByVal y As Double_** , **_ByVal CancelDefault As Boolean_** )

 _expression_ A variable that represents an **InvisibleApp** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Button_|Required| **Long**|The mouse button that was clicked. See Remarks for possible values.|
| _KeyButtonState_|Required| **Long**|The state of the mouse buttons and the SHIFT and CTRL keys for the event. See Remarks for possible values.|
| _x_|Required| **Double**|The x-coordinate of the mouse pointer.|
| _y_|Required| **Double**|The y-coordinate of the mouse pointer.|
| _CancelDefault_|Required| **Boolean**| **False** if Microsoft Visio should process the message it receives from this event; otherwise, **True** .|

## Remarks

Possible values for  _Button_ are shown in the following table, and are declared in **VisKeyButtonFlags** in the Visio type library.



|**Constant **|**Value **|
|:-----|:-----|
| **visMouseLeft**|1|
| **visMouseMiddle**|16|
| **visMouseRight**|2|
Possible values for  _KeyButtonState_ can be a combination of the values shown in the following table, which are declared in **VisKeyButtonFlags** in the Visio type library. For example, if _KeyButtonState_ returns 9, it indicates that the user clicked the left mouse button while pressing CTRL.



|**Constant **|**Value **|
|:-----|:-----|
| **visKeyControl**|8|
| **visKeyShift**|4|
| **visMouseLeft**|1|
| **visMouseMiddle**|16|
| **visMouseRight**|2|
If you set  _CancelDefault_ to **True** , Visio will not process the message received when the mouse button is clicked.

Unlike some other Visio events,  **MouseMove** does not have the prefix "Query," but it is nevertheless a query event. That is, you can cancel processing the message sent by **MouseMove** , either by setting _CancelDefault_ to **True** , or, if you are using the **VisEventProc** method to handle the event, by returning **True** . For more information, see the topics for the **VisEventProc** method and for any of the query events (for example, the **QueryCancelSuspend** event) in this Automation Reference.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


## Example

This class module shows how to define a sink class called  **MouseListener** that listens for events fired by mouse actions in the active window. It declares the object variable _vsoWindow_ by using the **WithEvents** keyword. The class module also contains event handlers for the **MouseDown** , **MouseMove** , and **MouseUp** events.

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
 
 If Button = 1 Then 
 
 Debug.Print "Left mouse button clicked" 
 
 ElseIf Button = 2 Then 
 
 Debug.Print "Right mouse button clicked" 
 
 ElseIf Button = 16 Then 
 
 Debug.Print "Center mouse button clicked" 
 
 End If 
 
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

Save the document to initialize the class, and then move the mouse briefly in the active window to fire a series of  **MouseMove** events. In the Immediate window, the handler prints a list of x- and y-position value pairs, each of which corresponds to the mouse position when a **MouseMove** event fired.


