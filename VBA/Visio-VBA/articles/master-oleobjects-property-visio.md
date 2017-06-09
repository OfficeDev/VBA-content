---
title: Master.OLEObjects Property (Visio)
keywords: vis_sdr.chm10713965
f1_keywords:
- vis_sdr.chm10713965
ms.prod: visio
api_name:
- Visio.Master.OLEObjects
ms.assetid: b51fbdc2-a236-4733-5a2e-b8e75d457d64
ms.date: 06/08/2017
---


# Master.OLEObjects Property (Visio)

Returns the  **OLEObjects** collection of a master. Read-only.


## Syntax

 _expression_ . **OLEObjects**

 _expression_ A variable that represents a **Master** object.


### Return Value

OLEObjects


## Remarks

The  **OLEObjects** property returns an **OLEObjects** collection that includes any OLE 2.0 linked or embedded objects or ActiveX controls contained in a document, master, or page.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to get the  **OLEObjects** collection of an active page and print the **ClassID** and **ProgID** for each **OLEObject** object in the Immediate window. This example assumes that the active page contains at least one OLE 2.0 embedded or linked object or an ActiveX control.


```vb
 
Public Sub OLEObjects_Example() 
 
 Dim intCounter As Integer 
 Dim vsoOLEObjects As Visio.OLEObjects 
 
 'Get the OLEObjects collection of the active page. 
 Set vsoOLEObjects = ActivePage.OLEObjects 
 
 'Step through the collection of OLEObjects on the page. 
 For intCounter = 1 To vsoOLEObjects.Count 
 Debug.Print vsoOLEObjects(intCounter).ClassID 
 Debug.Print vsoOLEObjects(intCounter).ProgID 
 Next intCounter 
 
End Sub
```


