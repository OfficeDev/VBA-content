---
title: Masters.Add Method (Visio)
keywords: vis_sdr.chm10816005
f1_keywords:
- vis_sdr.chm10816005
ms.prod: visio
api_name:
- Visio.Masters.Add
ms.assetid: 3951e242-c7e6-7a30-bf2c-0af7c030ace1
ms.date: 06/08/2017
---


# Masters.Add Method (Visio)

Adds a new object to a collection.


## Syntax

 _expression_ . **Add**

 _expression_ A variable that represents a **Masters** object.


### Return Value

Master


## Remarks

All properties of the new object are initialized to zero, so you need to set only the properties that you want to change.


## Example

The following macro shows how to add  **Master** objects to the **Masters** collection and **Page** objects to the **Pages** collection. It also shows how to add documents, layers, styles, events, and add-ons to their corresponding collections.

Before running this macro, replace  _Myfile.vsd_ with a valid .vsd file and references to _path_ \ _filename_ and _filename_ with a valid path and/or file name to an execuatable add-on (EXE) in your Visio project. The add-on should take no arguments.




```vb
Public Sub Add_Example() 
 
 Dim vsoMasters As Visio.Masters 
 Dim vsoAddons As Visio.Addons 
 Dim vsoPages As Visio.Pages 
 Dim vsoEventList As Visio.EventList 
 Dim vsoLayers As Visio.Layers 
 Dim vsoLayer As Visio.Layer 
 Dim vsoStyles As Visio.Styles 
 
 Dim vsoDocument As Visio.Document 
 Dim vsoEvent As Visio.Event 
 Dim vsoMaster As Visio.Master 
 Dim vsoPage As Visio.Page 
 Dim vsoShape As Visio.Shape 
 Dim vsoStyle As Visio.Style 
 Dim vsoAddon As Visio.Addon 
 
 'Add a document based on the Basic Diagram template. 
 Set vsoDocument = Documents.Add("Basic Diagram.vst") 
 
 'Add a document based on a drawing (creates a copy of the drawing). 
 Set vsoDocument = Documents.Add("Myfile.vsd ") 
 
 'Add a document based on a stencil (creates a copy of the stencil). 
 Set vsoDocument = Documents.Add("Basic Shapes.vss") 
 
 'Add a document object based on no template. 
 Set vsoDocument = Documents.Add("") 
 
 'Get the Pages collection and add a page to the collection. 
 Set vsoPages = vsoDocument.Pages 
 Set vsoPage = vsoPages.Add 
 
 'Get the Masters collection and add a master to the collection. 
 Set vsoMasters = vsoDocument.Masters 
 Set vsoMaster = vsoMasters.Add 
 
 'Get the Layers collection and add a layer named "MyLayer" 
 'to the collection. 
 Set vsoLayers = vsoPage.Layers 
 Set vsoLayer = vsoLayers.Add("MyLayer") 
 
 'Draw a rectangle. 
 Set vsoShape = vsoPage.DrawRectangle(3, 3, 5, 6) 
 
 'Add this shape to MyLayer. The second argument is required but has 
 'no effect, because vsoShape is not a group shape. 
 vsoLayer.Add vsoShape, 0 
 
 'Add a style named "My FillStyle" to the Styles collection. 
 'This style is based on the Basic style and includes 
 'only a Fill style. 
 Set vsoStyles = vsoDocument.Styles 
 Set vsoStyle = vsoStyles.Add("My FillStyle", "Basic", False, False, True) 
 
 'Add a style named "My NoStyle" to the Styles collection. 
 'This style is based on no style and includes 
 'Text, Line, and Fill styles. 
 Set vsoStyle = vsoStyles.Add("My NoStyle", "", True, True, True) 
 
 'Add an add-on to the Addons collection. 
 Set vsoAddons = Visio.Addons 
 Set vsoAddon = vsoAddons.Add("path \filename ") 
 
 'Add a BeforeDeleteSelection event to the EventList collection 
 'of the document. The event will start your add-on, which 
 'takes no arguments. 
 Set vsoEventList = vsoDocument.EventList 
 Set vsoEvent = vsoEventList.Add(visEvtCodeBefSelDel, visActCodeRunAddon, _ 
 "filename ", "") 
 
End Sub
```


