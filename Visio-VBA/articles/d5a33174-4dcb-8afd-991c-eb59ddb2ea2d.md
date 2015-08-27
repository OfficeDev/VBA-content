
# IVisEventProc.VisEventProc Method (Visio)

 **Last modified:** July 28, 2015

 _**Applies to:** Visio 2013 Preview_

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Private member function of  **IVisEventProc** that handles event notifications passed to it by the **EventList.AddAdvise** method.

## Syntax
<a name="sectionSection0"> </a>

 _expression_. **VisEventProc**( **_nEventCode_**,  **_pSourceObj_**,  **_nEventID_**,  **_nEventSeqNum_**,  **_pSubjectObj_**,  **_vMoreInfo_**)

 _expression_A variable that represents an  **IVisEventProc** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|nEventCode|Required| **Integer**|The event or events that occurred. |
|pSourceObj|Required| **Object**|The object whose  **EventList** collection contains the **Event** object that sent the notification.|
|nEventID|Required| **Long**|The unique identifier of the  **Event** object within the **EventList** collection.|
|nEventSeqNum|Required| **Long**|The ordinal position of the event with respect to the sequence of events that have occurred in the calling instance of the application. |
|pSubjectObj|Required| **Object**|The subject of the event, which is the object to which the event occurred. See Remarks for examples.|
|vMoreInfo|Required| **Variant**|Additional information about the subject of the event. See Remarks for more information.|

### Return Value

Variant


## Remarks
<a name="sectionSection1"> </a>

To handle event notifications, create a class module that implements the  **IVisEventProc** interface and then create an instance of this class to pass as an argument to the **AddAdvise** method of the **EventList** collection. Use the **AddAdvise** method to create **Event** objects that send the notifications.

The nEventCode parameter identifes the specific event or events that occurred. TheEventCode argument of the **AddAdvise** method is passed to **VisEventProc** asnEventCode. Within your procedure, you can use any branching technique you want to determine which event occurred and handle it. The example that accompanies this topic uses a  **Select Case** decision structure.

Unlike the  **Index** property of the **EventList** collection,nEventID does not change as **Event** objects are added or deleted from the collection.

From within  **VisEventProc** , you can use the following code to get the **Event** object that sent the notification:




```
pSourceObj.EventList.ItemFromID(nEventID)
```

The connection between the source object pSourceObj and the **Event** object exists until one of the following occurs:


- The program deletes the  **Event** object.
    
- The program releases the last reference to the source object. (The  **EventList** collection and **Event** objects hold a reference to their source object.)
    
- The Microsoft Visio application instance terminates.
    
The first event that occurs in a Visio instance has nEventSeqNum = 1, the second event = 2, and so on. In some cases, you can use the sequence number in conjunction with the **EventInfo** property to obtain more information about the event.

ThepSubjectObj parameter for a **ShapeAdded** event is a **Shape** object that represents the shape that was just added, while the subject of a **BeforeSelectionDelete** event is a **Selection** object in which the shapes that are about to be deleted are selected.

For many events, vMoreInfo is a string similar to the command line the application passes to the add-ons it executes. If the notification does not include additional information, this parameter is set to **Nothing**. For details about notification parameters for a particular event, see the particular event topic in this Automation Reference.

 Beginning with Visio 2000, **VisEventProc** is defined as a function that returns a value. However, Visio only looks at return values from calls to **VisEventProc** that are passed a query event code. Sink objects that provide **VisEventProc** through **IDispatch** require no change. To modify existing event handlers so that they can handle query events, change the **Sub** procedure to a **Function** procedure and return the appropriate value. (For details about query events, see this reference for event topics prefixed with **Query**.)

If nEventCode identifies a query event (events prefixed with **Query**), return  **True** from **VisEventProc** to cancel the event, and return **False** to allow it to happen. The value is arbitrary for other events. If you do not return an explicit value, Microsoft Visual Basic for Applications (VBA) returns an empty **Variant**, which Visio interprets as  **False**.


## Example
<a name="sectionSection2"> </a>

This example shows how to create a class module that implements  **IVisEventProc** to handle events fired by a source object in Visio, for example, the **Document** object. The module consists of the function **VisEventProc** , which uses a **Select Case** block to check for three events: **DocumentSaved**,  **PageAdded**, and  **ShapesDeleted**. Other events fall under the default case ( **Case Else**). Each  **Case** block constructs a string (strMessage) that contains the name and event code of the event that fired. Finally, the function displays the string in the Immediate window.

Copy this sample code into a new class module in VBA or Visual Basic, naming the module  **clsEventSink**. You can then use an event-sink module to create an instance of the  **clsEventSink** class and **Event** objects that send notifications of event firings to the class instance. To see how to create an event-sink module, see the example for the **AddAdvise** method.




```
 
Implements Visio.IVisEventProc 
 
 
'Declare visEvtAdd as a 2-byte value 
'to avoid a run-time overflow error 
Private Const visEvtAdd% = &amp;H8000 
 
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
 strMessage = "DocumentSaved (" &amp; nEventCode &amp; ")" 
 Case (visEvtPage + visEvtAdd) 
 strMessage = "PageAdded (" &amp; nEventCode &amp; ")" 
 Case visEvtCodeShapeDelete 
 strMessage = "ShapesDeleted(" &amp; nEventCode &amp; ")" 
 Case Else 
 strMessage = "Other (" &amp; nEventCode &amp; ")" 
 End Select 
 
 
 'Display the event name and the event code 
 Debug.Print strMessage 
 
End Function 

```

