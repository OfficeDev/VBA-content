---
title: KeyboardEvent.ToString Property (Visio)
keywords: vis_sdr.chm17051505
f1_keywords:
- vis_sdr.chm17051505
ms.prod: visio
api_name:
- Visio.KeyboardEvent.ToString
ms.assetid: 039e4d80-dcff-0781-5ae4-0bc2a9b7a6d8
ms.date: 06/08/2017
---


# KeyboardEvent.ToString Property (Visio)

Returns a string that represents the properties of a  **KeyboardEvent** or **MouseEvent** object. Read-only.


## Syntax

 _expression_ . **ToString**

 _expression_ A variable that represents a **KeyboardEvent** object.


### Return Value

String


## Remarks

 **ToString** is the default property of both **KeyboardEvent** and **MouseEvent** objects.

When a  **KeyDown** , **KeyPress** , or **KeyUp** event fires, the **ToString** property returns a string that represents the properties of the **KeyboardEvent** object that gets passed to **VisEventProc** . The string has the following form:

 _event code_ ; **KeyCode** property value; **KeyButtonState** property value; **KeyAscii** property value; **Window.Caption**

where  _event code_ returns the code of the event that fired and **Window.Caption** returns the caption of the window that sourced the event. For example, if a user pressed the "L" key while holding down the SHIFT key, in response to the **KeyPress** event, **ToString** might return

713;0;4;76;Drawing1

When a  **MouseDown** , **MouseMove** , or **MouseUp** event fires, the **ToString** property returns a string that represents the properties of the **MouseEvent** object that gets passed to **VisEventProc** . The string has the following form:

 _event code_ ; **Button** property value; **KeyButtonState** property value; **x** property value; **y** property value; **Window.Caption**

where  _event code_ returns the code of the event that fired and **Window.Caption** returns the caption of the window that sourced the event. For example, if a user clicked the left mouse button near the middle of the drawing page while holding down the SHIFT key, in response to the **MouseDown** event, **ToString** might return

709;1;5;4.3750003+000;4.265000+000;Drawing1

For more information about the possible values returned by each of the individual properties represented by the string returned by  **ToString** , see the respective property topics in this Automation Reference.


## Example

The following Microsoft Visual Basic for Applications (VBA) example shows how to use the  **AddAdvise** method to create a **Event** object that will sink a **MouseDown** event. It uses the **ToString** property of the **MouseEvent** object to report the details of the event that fired.

The example contains a class module and two public procedures that are inserted into the  **ThisDocument** project of the active Visio document:




- The  **CreateEventObject** procedure creates an instance of a sink-object (event-handling) class named **clsEventSink** that gets passed to the **AddAdvise** method, and that receives notifications of events. In addition, the procedure creates an **Event** object to send notifications of firings of the **MouseDown** event sourced by the **Application** object to the sink object.
    
- The  **DeleteEventObject** procedure deletes this **Event** object when your program is finished using it.
    


The  **clsEventSink** class implements the **IVisEventProc** interface. The class module creates a class to handle events fired by the Visio **Application** object. The module consists of the function **VisEventProc** , which uses a **Select Case** block to check for the **MouseDown** event. When a **MouseDown** event fires, Visio passes a **MouseEvent** object to **VisEventProc** as _pSubjectObj_. The function then constructs a message that displays the string returned by the  **ToString** property of the **MouseEvent** object passed to the function.

Other events fall under the default case ( **Case Else** ). The **Case Else** block constructs a string ( _strMessage_ ) that contains the name and event code of the event that fired. Finally, the function displays the string in the Immediate window.

The example assumes that there is an active document in the Visio application window. Copy the following code into the  **ThisDocument** project in the Visual Basic Editor:




```vb
Option Explicit 
 
Private mEventSink As clsEventSink 
 
'Declare visEvtAdd as a 2-byte value 
'to avoid a run-time overflow error 
Private Const visEvtAdd% = &;H8000 
 
Public Sub CreateMouseDownEventObject() 
 
 Dim vsoApplicationEvents As Visio.EventList 
 Dim vsoMouseDownEvent As Visio.Event 
 
 'Create an instance of the clsEventSink class 
 'to pass to the AddAdvise method. 
 Set mEventSink = New clsEventSink 
 
 'Get the EventList collection of the application 
 Set vsoApplicationEvents = Application.EventList 
 
 'Add an Event object for the MouseDown event 
 'that will send notifications. 
 Set vsoMouseDownEvent= vsoApplicationEvents.AddAdvise( _ 
 visEvtCodeMouseDown, mEventSink, "", "Mouse down...") 
 
End Sub 
 
Public Sub DeleteMouseDownEventObject() 
 
 'Delete the Event object for the MouseDown event 
 vsoMouseDownEvent.Delete 
 Set vsoMouseDownEvent = Nothing 
 
End Sub
```

Copy the following code into a new class module in VBA, naming the module  **clsEventSink**. 




```vb
Implements Visio.IVisEventProc 
 
Private Function IVisEventProc_VisEventProc( _ 
 ByVal nEventCode As Integer, _ 
 ByVal pSourceObj As Object, _ 
 ByVal nEventID As Long, _ 
 ByVal nEventSeqNum As Long, _ 
 ByVal pSubjectObj As Object, _ 
 ByVal vMoreInfo As Variant) As Variant 
 
 Dim strMessage As String 
 Dim vsoMouseDownEvent As Visio.MouseEvent 
 
 'Find out which event fired 
 Select Case nEventCode 
 Case visEvtCodeMouseDown 
 Set vsoMouseEvent = pSubjectObj 
 strMessage = "ToString is: " &; vsoMouseEvent.ToString 
 Case Else 
 strMessage = "Other (" &; nEventCode &; ")" 
 End Select 
 
 'Display the event name and the event code 
 Debug.Print strMessage 
 
End Function
```


