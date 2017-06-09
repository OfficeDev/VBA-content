---
title: OLEObject.ClassID Property (Visio)
keywords: vis_sdr.chm15213240
f1_keywords:
- vis_sdr.chm15213240
ms.prod: visio
api_name:
- Visio.OLEObject.ClassID
ms.assetid: 9241135d-6c02-046b-02b4-f8d4b308878d
ms.date: 06/08/2017
---


# OLEObject.ClassID Property (Visio)

Returns the class ID string of a shape that represents an ActiveX control or an embedded or linked OLE object. Read-only.


## Syntax

 _expression_ . **ClassID**

 _expression_ A variable that represents an **OLEObject** object.


### Return Value

String


## Remarks

The  **ClassID** property raises an exception if the shape doesn't represent an ActiveX control or an OLE 2.0 embedded or linked object. A shape represents an ActiveX control or an OLE 2.0 embedded or linked object if the **visTypeIsOLE2** bit (&;H8000) is set in the value returned by **Shape** . **ForeignType** .

 **ClassID** returns a string of the form:




```
{2287DC42-B167-11CE-88E9-002AFDDD917}
```

This identifies the application that services the object. It might, for example, identify an embedded object on a Microsoft Visio page as a Microsoft Excel object.

After using a shape's  **Object** property to obtain an Automation interface on the object the shape represents, you might want to obtain the shape's **ClassID** or **ProgID** property to determine the methods and properties provided by the interface.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to get the  **OLEObjects** collection of an active page and print the **ClassID** for each **OLEObject** object in the Immediate window. This example assumes that the active page has at least one OLE 2.0 embedded or linked object or an ActiveX control.


```vb
 
Public Sub ClassID_Example() 
 
 Dim intCounter As Integer 
 Dim vsoOLEObjects As Visio.OLEObjects 
 
 'Get the OLEObjects collection of the active page. 
 Set vsoOLEObjects = ActivePage.OLEObjects 
 
 'Step through the collection of OLEObjects on the page. 
 For intCounter = 1 To vsoOLEObjects.Count 
 Debug.Print vsoOLEObjects(intCounter).ClassID 
 Next intCounter 
 
End Sub
```


