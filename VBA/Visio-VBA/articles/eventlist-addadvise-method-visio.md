---
title: EventList.AddAdvise Method (Visio)
keywords: vis_sdr.chm12716010
f1_keywords:
- vis_sdr.chm12716010
ms.prod: visio
api_name:
- Visio.EventList.AddAdvise
ms.assetid: b58e086f-59d2-9e63-5df3-3001b58bb2c1
ms.date: 06/08/2017
---


# EventList.AddAdvise Method (Visio)

Adds an  **Event** object to the **EventList** collection of the source object whose events you want to receive. When selected events occur, the source object notifies your sink object.


## Syntax

 _expression_ . **AddAdvise**( **_EventCode_** , **_SinkIUnkOrIDisp_** , **_IIDSink_** , **_TargetArgs_** )

 _expression_ A variable that represents an **EventList** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _EventCode_|Required| **Integer**|The event(s) that generate notifications.|
| _SinkIUnkOrIDisp_|Required| **Variant**|A reference to a COM interface on the object that is to receive event notifications.|
| _IIDSink_|Required| **String**| Reserved for future use. Must be "".|
| _TargetArgs_|Required| **String**|The string that is passed to your  **Event** object to set its **TargetArgs** property.|

### Return Value

Event


## Remarks

 **Event** objects created with the **AddAdvise** method have an **Action** property of **visActCodeAdvise** . They are not persistent, that is, they cannot be stored with a Visio document and must be re-created at run time.

The source object whose  **EventList** collection contains the **Event** object establishes the scope in which the events are reported. Events are reported for the source object and objects lower in the object model hierarchy. For example, to receive notification when a particular document is saved, add an **Event** object for the **DocumentSaved** event to the **EventList** collection of that document. To receive notification when any document is opened in an instance of the application, add the **Event** object to the **EventList** collection of the **Application** object.

Creating  **Event** objects is a common way to handle events from C++ or other non-Microsoft Visual Basic solutions. When you use the Visual Basic **WithEvents** keyword to handle events, all the events in a source object's event set fire. When you create **Event** objects to handle events, however, your program will only be notified of the events you select. Depending on your solution, this may result in improved performance.

The  _EventCode_ argument is often a combination of constants. For example, **visEvtMod** + **visEvtCell** is the event code for the **CellChanged** event. Event constants are declared by the Visio type library and are prefixed with **visEvt** . To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx). 

The arguments passed to the  **AddAdvise** method set the initial values of the **Event** object's **Event** , **Action** ( **visCodeRunAddAdvise** ), and **TargetArgs** properties.

Beginning with Visio 2002, you can use event filters to refine the events that you receive in your program. You can filter events by object, cell, ranges of cells, or command ID. For details about using event filters, see the method topics prefixed with  **SetFilter** and **GetFilter** .

 **Event** objects created with the **AddAdvise** method have an **Action** property of **visActCodeAdvise** . They are not persistent, that is, they cannot be stored with a Visio document and must be re-created at run time.

The source object whose  **EventList** collection contains the **Event** object establishes the scope in which the events are reported. Events are reported for the source object and objects lower in the object model hierarchy. For example, to receive notification when a particular document is saved, add an **Event** object for the **DocumentSaved** event to the **EventList** collection of that document. To receive notification when any document is opened in an instance of the application, add the **Event** object to the **EventList** collection of the **Application** object.

Creating  **Event** objects is a common way to handle events from C++ or other non-Microsoft Visual Basic solutions. When you use the Visual Basic **WithEvents** keyword to handle events, all the events in a source object's event set fire. When you create **Event** objects to handle events, however, your program will only be notified of the events you select. Depending on your solution, this may result in improved performance.

The  _EventCode_ argument is often a combination of constants. For example, **visEvtMod** + **visEvtCell** is the event code for the **CellChanged** event. Event constants are declared by the Visio type library and are prefixed with **visEvt** . To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx). 

The arguments passed to the  **AddAdvise** method set the initial values of the **Event** object's **Event** , **Action** ( **visCodeRunAddAdvise** ), and **TargetArgs** properties.

Beginning with Visio 2002, you can use event filters to refine the events that you receive in your program. You can filter events by object, cell, ranges of cells, or command ID. For details about using event filters, see the method topics prefixed with  **SetFilter** and **GetFilter** .


### Enabling your program to handle event notifications from Microsoft Visual Basic or Visual Basic for Applications

To handle notifications, you should create a class module that implements the  **IVisEventProc** interface and then create an instance of this class to pass as an argument to the **AddAdvise** method.

The  **IVisEventProc** interface contains a single function with the following declaration:




```vb
Implements IVisEventProc 
 
Private Function IVisEventProc_VisEventProc( _  
    ByVal nEventCode As Integer, _  
    ByVal pSourceObj As Object, _  
    ByVal nEventID As Long, _  
    ByVal nEventSeqNum As Long, _  
    ByVal pSubjectObj As Object, _  
    ByVal vMoreInfo As Variant) As Variant  
End Function
```

Following are descriptions of the arguments to  **VisEventProc** .



|** Argument**|** Description**|
|:-----|:-----|
| nEventCode| The event(s) that occurred. You can provide a distinct object for each event or provide a single object that receives all notifications and switches internally based on nEventCode.|
| pSourceObj| The object whose ** EventList** collection contains the **Event** object that triggered the notification.|
| nEventID|The unique identifier of the  **Event** object within the **EventList** collection. (Unlike the **Index** property of the **EventList** collection, nEventID does not change as **Event** objects are added to or deleted from the collection.) You can get the **Event** object from **VisEventProc** by using the following code:
```
pSourceObj .EventList.ItemFromID(nEventID )
```

|
| nEventSeqNum| The ordinal position of the event with respect to the sequence of events that have occurred in the calling instance of the application. The first event that occurs in a Visio instance has a sequence number of 1, the second event 2, and so forth. In some cases, you can use the sequence number in conjunction with the **EventInfo** property to obtain more information about the event.|
| pSubjectObj| The object that the event is about. For example, the subject of a **ShapeAdded** event is a **Shape** object representing the shape that was just added, while the subject of a **BeforeSelectionDelete** event is a **Selection** object in which the shapes that are about to be deleted are selected.|
| vMoreInfo| Additional information about the subject of the event. For many events, it is a string similar to the command line the application passes the add-ons it executes. If the notification does not include additional information, this parameter is set to **Nothing** . For details about notification parameters for a particular event, see the particular event topic in this Automation Reference.|
If nEventCode identifies a query event (events prefixed with  **Query** ), return **True** to cancel the event, and **False** to allow it to happen. The value is arbitrary for other events. If no explicit value is returned, Microsoft Visual Basic for Applications (VBA) returns an empty **Variant** , which Visio interprets as **False** .

The connection between the source object and the  **Event** object exists until one of the following occurs:


- The program deletes the  **Event** object.
    
- The program releases the last reference to the source object. (The  **EventList** collection and **Event** objects hold a reference on their source object.)
    
- The application terminates.
    
 Beginning with Visio 2000, **VisEventProc** is defined as a function that returns a value. However, Visio only looks at return values from calls to **VisEventProc** that are passed a query event code. **Event** objects that provide **VisEventProc** through **IDispatch** require no change. To modify existing event handlers so that they can handle query events, change the **Sub** procedure to a **Function** procedure and return the appropriate value. (For details about query events, see this reference for event topics prefixed with **Query** .)


## Example

This example shows how to create a class module to handle events fired by a source object in Microsoft Office Visio, for example, the  **Document** object. The module consists of the function **VisEventProc** , which uses a **Select Case** block to check for three events: **DocumentSaved** , **PageAdded** , and **ShapesDeleted** . Other events fall under the default case ( **Case Else** ). Each **Case** block constructs a string ( _strMessage_ ) that contains the name and event code of the event that fired. Finally, the function displays the string in the Immediate window.

Copy this sample code into a new class module in Microsoft Visual Basic for Applications (VBA) or Visual Basic, naming the module  **clsEventSink** . You can then use the event-sink module that follows to create an instance of the **clsEventSink** class and **Event** objects that send notifications of event firings to the class instance.




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
     
    'Find out which event fired 
    Select Case nEventCode  
        Case visEvtCodeDocSave  
            strMessage = "DocumentSaved (" &; nEventCode &; ")"  
        Case (visEvtPage + visEvtAdd)  
            strMessage = "PageAdded (" &; nEventCode &; ")"  
        Case visEvtCodeShapeDelete 
            strMessage = "ShapesDeleted(" &; nEventCode &; ")"  
        Case Else  
            strMessage = "Other (" &; nEventCode &; ")"  
    End Select 
     
    'Display the event name and the event code 
    Debug.Print strMessage  
 
End Function
```

The following VBA module shows how to use the  **AddAdvise** method to sink events. The module contains two public procedures.

The  **CreateEventObjects** procedure creates an instance of a sink-object (event-handling) class named **clsEventSink** that gets passed to the **AddAdvise** method, and that receives notifications of events. In addition, the procedure creates three **Event** objects to send notifications of firings of three events sourced by the **Document** object to the sink object: **DocumentSaved** , **PageAdded** , and **ShapesDeleted** .



The  **DeleteEventObjects** procedure deletes these **Event** objects when your program is finished using them.

The  **clsEventSink** class implements the **IVisEventProc** interface.

The example assumes that there is an active document in the Visio application window.






```vb
Option Explicit 
 
Private mEventSink As clsEventSink 
 
Dim vsoDocumentEvents As Visio.EventList       
Dim vsoDocumentSavedEvent As Visio.Event  
Dim vsoPageAddedEvent As Visio.Event  
Dim vsoShapesDeletedEvent As Visio.Event 
    
'Declare visEvtAdd as a 2-byte value 
'to avoid a run-time overflow error 
Private Const visEvtAdd% = &;H8000  
 
Public Sub CreateEventObjects()      
 
    'Create an instance of the clsEventSink class 
    'to pass to the AddAdvise method. 
    Set mEventSink = New clsEventSink 
  
    'Get the EventList collection of the active document. 
    Set vsoDocumentEvents = ActiveDocument.EventList  
 
    'Add Event objects that will send notifications. 
    'Add an Event object for the DocumentSaved event. 
    Set vsoDocumentSavedEvent= vsoDocumentEvents.AddAdvise( _  
     visEvtCodeDocSave, mEventSink, "", "Document saved...")  
 
    'Add an Event object for the PageAdded event. 
    Set vsoPageAddedEvent= vsoDocumentEvents.AddAdvise( _  
     visEvtAdd + visEvtPage, mEventSink, "", "Page added...")  
 
    'Add an Event object for the ShapesDeleted event. 
    Set vsoShapesDeletedEvent = vsoDocumentEvents.AddAdvise( _  
     visEvtCodeShapeDelete, mEventSink, "", "Shapes deleted...")  
 
End Sub   
 
Public Sub DeleteEventObjects()  
 
    'Delete the Event object for the DocumentSaved event.    
    vsoDocumentSavedEvent.Delete  
    Set vsoDocumentSavedEvent = Nothing 
 
    'Delete the Event object for the PageAdded event. 
    vsoPageAddedEvent.Delete  
    Set vsoPageAddedEvent = Nothing 
 
    'Delete the Event object for the ShapesDeleted event.  
    vsoShapesDeletedEvent.Delete  
    Set vsoShapesDeletedEvent = Nothing 
 
End Sub
```


