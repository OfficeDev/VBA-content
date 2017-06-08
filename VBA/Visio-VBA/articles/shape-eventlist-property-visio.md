---
title: Shape.EventList Property (Visio)
keywords: vis_sdr.chm11213480
f1_keywords:
- vis_sdr.chm11213480
ms.prod: visio
api_name:
- Visio.Shape.EventList
ms.assetid: 513838c2-f00e-06e3-f08b-b23295f7f0d3
ms.date: 06/08/2017
---


# Shape.EventList Property (Visio)

Returns the  **EventList** collection of an object or the **EventList** collection that contains an **Event** object. Read-only.


## Syntax

 _expression_ . **EventList**

 _expression_ A variable that represents a **Shape** object.


### Return Value

EventList


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **EventList** property to add an **Event** object to the **EventList** collection of a **Document** object. When the **Event** object is triggered by adding a shape to the document, the VSL add-on you specify runs.

Before running this macro, replace references to  _fullpath\filename_ and _filename_ with a valid path and name for a Microsoft Visio VSL or executable (EXE) add-on.




```vb
 
Public Sub EventList_Example() 
 
 Dim vsoEventList As Visio.EventList 
 Dim vsoEvent As Visio.Event 
 Dim vsoAddons As Visio.Addons 
 Dim vsoAddon As Visio.Addon 
 
 'Prevent overflow error. 
 Const visEvtAdd% = &;H8000 
 
 'Add the specified add-on to the Addons collection. 
 Set vsoAddons = Visio.Addons 
 Set vsoAddon = vsoAddons.Add ("fullpath\filename ") 
 
 'Add a ShapeAdded event to the EventList collection 
 'of the document. The event will start the specifed add-on, which 
 'should take no arguments. 
 Set vsoEventList = ThisDocument.EventList 
 Set vsoEvent = vsoEventList.Add(visEvtAdd + visEvtShape, visActCodeRunAddon, _ 
 "filename ", "") 
 
End Sub
```


