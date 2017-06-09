---
title: MouseEvent.DragState Property (Visio)
keywords: vis_sdr.chm17160265
f1_keywords:
- vis_sdr.chm17160265
ms.prod: visio
api_name:
- Visio.MouseEvent.DragState
ms.assetid: 958fa39f-5ca4-3911-72f5-2bea3c1ded48
ms.date: 06/08/2017
---


# MouseEvent.DragState Property (Visio)

Returns information about the state of mouse movement as it relates to dragging and dropping a shape. Read-only.


## Syntax

 _expression_ . **DragState**

 _expression_ An expression that returns a **MouseEvent** object.


### Return Value

Long


## Remarks

The  **DragState** property extends the **MouseMove** event by returning detailed information about the state of mouse movements and actions throughout the course of a drag and drop operation. You can use the **DragState** property in conjunction with the **[EventList.AddAdvise ](eventlist-addadvise-method-visio.md)** method to determine whether a drag and drop operation is beginning, or whether the mouse is entering a drop-target window, moving over the window, dropping an object in the target window, or leaving the window.


 **Note**  You can specify exactly which drag-states extensions you want to listen to by using the  **[Event.SetFilterActions](event-setfilteractions-method-visio.md)** method.

To handle event notifications, create a class module that implements the  **[VisEventProc](iviseventproc-viseventproc-method-visio.md)** method of the **IVisEventProc** interface and then create an instance of this class to pass as an argument to the **AddAdvise** method. Get the value of the **DragState** property of the _pSubjectObj_ parameter of the **VisEventProc** function.

At any time you can return  **VisEventProc** = **True** to cancel the drag and drop action, for example if you receive an event notification that the user is attempting to drop an object in an inappropriate target window.

The example that accompanies this topic provides sample code that shows how to get drag-state information.

Possible values returned by the  **DragState** property are shown in the following table and declared in the **VisMouseMoveDragStates** enumeration, which is declared in the Microsoft Visio type library.



|**Constant**|**Value **|**Description**|
|:-----|:-----|:-----|
| **visMouseMoveDragStatesBegin**|1|User is beginning to drag an object with the mouse.|
| **visMouseMoveDragStatesDrop**|5|User has dropped the dragged object in the drop-target window.|
| **visMouseMoveDragStatesEnter**|2|User is dragging an object into the drop-target window with the mouse.|
| **visMouseMoveDragStatesLeave**|4|User is moving the mouse out of the drop-target window.|
| **visMouseMoveDragStatesNone**|0|Either not a mouse movement or a mouse movement that is not a drag action.|
| **visMouseMoveDragStatesOver**|3|User is moving the dragged object within the drop-target window with the mouse.|
When the  **DragState** property returns **visMouseMoveDragStatesBegin** a drag and drop action is beginning. The **DragState** property returns **visMouseMoveDragStatesBegin** just once for each drag and drop action. At this point, you can cancel the drag and drop action entirely; if you do so, Visio fires no additional **MouseMove** events for any target windows.

When the  **DragState** property returns **visMouseMoveDragStatesEnter** , an end-user is dragging an object into a drop-target window. This event is fired once per drop-target window. At this point, you can cancel the drag and drop action for that specific drop-target window.

When the  **DragState** property returns **visMouseMoveDragStatesOver** , the user is dragging an object over a drop-target window. You can cancel the drag action, based on the type of window or on an _x,y_ range within a window, as specified in your code. Canceling a drag action over the drop-target window prevents the end-user from completing the drag and drop action.

When the  **DragState** property returns **visMouseMoveDragStatesDrop** , the drop-target window is receiving a drop. You can cancel the drop action, thus preventing the drop from occurring. When this occurs and you do not also cancel the drag action over the drop-target window, the end-user does not get any visual feedback to indicate that the drop action has been prevented.

When the  **DragState** property returns **visMouseMoveDragStatesLeave** , the end-user is moving the mouse out of the drop-target window. There is no way for you to cancel this operation at this point, but there would also be no logical reason to do so.


## Example

This example shows how to create a class module that implements the  **IVisEventProc** interface to handle events fired by the **MouseEvent** object. The module consists of the function **VisEventProc** , which uses a **Select Case** block to determine if the event that fired was a **MouseMove** event. If so, the code uses an **If...Else** block and the **DragState** property to determine the particular **MouseMove** event extension that fired.

Events other than  **MouseMove** fall under the default case ( **Case Else** ). In both cases, the **Case** block constructs a string ( _strMessage_ ) that contains the name and event code of the event that fired, including the drag-state extension and the _x-_ and _y-_ values of the location where the event fired, derived from the values of the **MouseEvent.X** and **MouseEvent.Y** properties. Finally, the function displays the message in the **Immediate** window.

 Copy this sample code into a new class module in VBA or Visual Basic, naming the module **clsEventSink** . You can then use the event-sink module that follows to create an instance of the **clsEventSink** class and an **Event** object for the **MouseMove** event that sends notifications of event firings to the class instance.




```vb
Implements Visio.IVisEventProc 
 
'Declare visEvtAdd as a 2-byte value 
'to avoid a run-time overflow error 
Private Const visEvtAdd% = &;H8000 
 
Private Function IVisEventProc_VisEventProc( _ 
    ByVal nEventCode As Integer, _ 
    ByVal pSourceObj As Object, _ 
    ByVal nEventID As Long, _ 
    ByVal nEventSeqNum As Long, _ 
    ByVal pSubjectObj As Object, _ 
    ByVal vMoreInfo As Variant) As Variant 
 
    Dim strMessage As String 
     
    'Find out which event and event extension fired 
    Select Case nEventCode 
        Case visEvtCodeMouseMove 
            Dim strInfo As String 
            If (pSubjectObj.DragState = visMouseMoveDragStatesOver) Then 
                strMessage = "MouseMove - dragOver (" + Str(pSubjectObj.x) + "," + Str(pSubjectObj.y) + ")" 
            ElseIf (pSubjectObj.DragState = visMouseMoveDragStatesBegin) Then 
               strMessage = "MouseMove - dragBegin (" + Str(pSubjectObj.x) + "," + Str(pSubjectObj.y) + ")" 
               If (pSubjectObj.Window.Index <> 1) Then 
                    IVisEventProc_VisEventProc = True       ' cancel for all windows except first one 
               End If 
            ElseIf (pSubjectObj.DragState = visMouseMoveDragStatesLeave) Then 
                strMessage = "MouseMove - dragLeave" 
            ElseIf (pSubjectObj.DragState = visMouseMoveDragStatesEnter) Then 
                strMessage = "MouseMove - dragEnter*******************************************" 
            ElseIf (pSubjectObj.DragState = visMouseMoveDragStatesDrop) Then 
                strMessage = "MouseMove - dragDrop" 
            End If 
        Case Else 
            strMessage = "Other (" &; nEventCode &; ")" 
    End Select 
     
    'Display the event name and the event code 
    If (Len(strMessage)) Then 
        Debug.Print strMessage 
    End If        
 
End Function
```

The following VBA module shows how to use the  **AddAdvise** method to sink events. The module contains two public procedures.

The  **CreateEventObjects** procedure creates an instance of a sink-object (event-handling) class named **clsEventSink** that gets passed to the **AddAdvise** method, and that receives notifications of events. In addition, the procedure creates a single **Event** object to send notifications of firings of the **MouseMove** event sourced by the **Application** object to the sink object.

The  **Initialize** procedure calls the **CreateEventObjects** procedure to start listening to events.

The  **clsEventSink** class implements the **IVisEventProc** interface.




```vb
Public Sub Initialize()     
 
    CreateEventObjects     
 
End Sub 
 
Option Explicit  
 
Private mEventSink As clsEventSink  
 
Dim vsoApplicationEvents As Visio.EventList  
Dim vsoMouseMoveEvent As Visio.Event    
 
'Declare visEvtAdd as a 2-byte value 
'to avoid a run-time overflow error 
Private Const visEvtAdd% = &;H8000  
 
Public Sub CreateEventObjects()  
 
    'Create an instance of the clsEventSink class 
    'to pass to the AddAdvise method. 
    Set mEventSink = New clsEventSink  
 
    'Get the EventList collection of the current instance of the Visio Application object 
    Set vsoApplicationEvents = Application.EventList  
 
   'Add an Event object that sends notifications of the MouseMove event. 
    Set vsoMouseMoveEvent = vsoApplicationEvents.AddAdvise(visEvtCodeMouseMove, mEventSink, "", "Mouse moved...")      
 
End Sub
```


