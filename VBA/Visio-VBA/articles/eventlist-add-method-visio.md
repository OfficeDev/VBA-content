---
title: EventList.Add Method (Visio)
keywords: vis_sdr.chm12716665
f1_keywords:
- vis_sdr.chm12716665
ms.prod: visio
api_name:
- Visio.EventList.Add
ms.assetid: 4deacea0-f4dc-bec6-31d0-bb9a060812f5
ms.date: 06/08/2017
---


# EventList.Add Method (Visio)

Adds an  **Event** object that runs an add-on when an event occurs. The **Event** object is added to the **EventList** collection of the source object whose events you want to receive.


## Syntax

 _expression_ . **Add**( **_EventCode_** , **_Action_** , **_Target_** , **_TargetArgs_** )

 _expression_ A variable that represents an **EventList** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _EventCode_|Required| **Integer**|The event(s) to capture.|
| _Action_|Required| **Integer**|The action to perform. Must be  **visActCodeRunAddon** , a member of **[VisEventCodes](viseventcodes-enumeration-visio.md)** in the Visio type Library.|
| _Target_|Required| **String**|The name of your add-on.|
| _TargetArgs_|Required| **String**|The string that is passed to your  **Event** object to set its **TargetArgs** property.|

### Return Value

Event


## Remarks

The source object whose  **EventList** collection contains the **Event** object establishes the scope in which the events are reported. Events are reported for the source object and objects lower in the object model hierarchy. For example, to run an add-on when a particular document is opened, add an **Event** object for the **DocumentOpened** event to the **EventList** collection of that document. To run an add-on when any document is opened in an instance of the application, add the **Event** object to the **EventList** collection of the **Application** object.

Creating  **Event** objects is a common way to handle events from C++ or other non-Microsoft Visual Basic solutions. When you use the Visual Basic **WithEvents** keyword to handle events, all the events in a source object's event set fire, but when you create **Event** objects, your program will only be notified of the events you select. Depending on your solution, this may result in improved performance.

 **Event** objects that run add-ons can be persistent: that is, they can be stored with a Visio document. To be persistent, an **Event** object's **Persistent** and **Persistable** properties must both be **True** .

The arguments passed to the  **Add** method set the initial values of the **Event** object's **Event** , **Action** ( **visActCodeRunAddon** ), **Target** , and **TargetArgs** properties.

Event codes are declared by the Visio type library and have the prefix  **visEvt** . Event codes are often a combination of constants. For example, **visEvtAdd** + **visEvtDoc** is the event code for the **DocumentAdded** event. To find an event code for the event you want to create, see[Event Codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).

To create an  **Event** object that advises the caller's sink object about an event, see the **AddAdvise** method.


## Example

The following example shows how to add an  **Event** object that runs an add-on to the **EventList** collection of the source object, in this case a **Document** object, whose events you want to receive.

Before running this macro, replace _path_ \ _filename_ with a valid path and file name for an executable add-on (EXE) in your Visio project. The add-on should take no arguments.




```vb
Public Sub AddEvent_Example() 
 
 Dim vsoAddons As Visio.Addons 
 Dim vsoEventList As Visio.EventList 
 Dim vsoDocument As Visio.Document 
 Dim vsoEvent As Visio.Event 
 Dim vsoAddon As Visio.Addon 
 
 'Add a document based on the Basic Diagram template. 
 Set vsoDocument = Documents.Add("Basic Diagram.vst") 
 
 'Add an add-on to the Addons collection. 
 Set vsoAddons = Visio.Addons 
 Set vsoAddon = vsoAddons.Add("path \filename ") 
 
 'Add a BeforeDeleteSelection event to the EventList collection 
 'of the Document object. The event will start your add-on, 
 'which takes no arguments. 
 Set vsoEventList = vsoDocument.EventList 
 Set vsoEvent = vsoEventList.Add(visEvtCodeBefSelDel, _ 
 visActCodeRunAddon, _ 
 "path \filename ", "") 
 
End Sub
```


