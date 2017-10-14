---
title: Shape.ProgID Property (Visio)
keywords: vis_sdr.chm11214160
f1_keywords:
- vis_sdr.chm11214160
ms.prod: visio
api_name:
- Visio.Shape.ProgID
ms.assetid: 2cd96dd5-7d73-77ea-9e7e-3d1dcd98a21a
ms.date: 06/08/2017
---


# Shape.ProgID Property (Visio)

Returns the programmatic identifier of a shape that represents an ActiveX control, an embedded object, or linked object. Read-only.


## Syntax

 _expression_ . **ProgID**

 _expression_ A variable that represents a **Shape** object.


### Return Value

String


## Remarks

The  **ProgID** property raises an exception if the shape doesn't represent an ActiveX control or OLE 2.0 embedded or linked object. A shape represents an ActiveX control, embedded object, or linked object if the **ForeignType** property returns **visTypeIsOLE2** in the value.

Use the  **ProgID** property of a **Shape** object or **OLEObject** to obtain the programmatic identifier of the object. Every OLE object class stores a programmatic identifier for itself in the registry. Typically this occurs when the program that services the object installs itself. Client programs use this identifier to identify the object. You are using the Microsoft Visio identifier when you execute a statement such as **GetObject** (,"Visio.Application") from a Microsoft Visual Basic program.

These are strings that the  **ProgID** property might return:




```
 
Visio.Drawing.5 
MSGraph.Chart.5 
Forms.CommandButton.1 

```

After using a shape's  **Object** property to obtain an **IDispatch** interface on the object the shape represents, you can obtain the shape's **ClassID** or **ProgID** property to determine the methods and properties provided by that interface.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to get the  **OLEObjects** collection of an active page and print the **ProgID** for each **OLEObject** object in the Immediate window. This example assumes that the active page has at least one OLE 2.0 embedded or linked object or an ActiveX control.


```vb
 
Public Sub ProgID_Example() 
 
 Dim intCounter As Integer 
 Dim vsoOLEObjects As Visio.OLEObjects 
 
 'Get the OLEObjects collection of the active page. 
 Set vsoOLEObjects = ActivePage.OLEObjects 
 
 'Step through the OLEObjects collection. 
 For intCounter = 1 To vsoOLEObjects.Count 
 Debug.Print vsoOLEObjects(intCounter).ProgID 
 Next intCounter 
 
End Sub
```


