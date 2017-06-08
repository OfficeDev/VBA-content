---
title: OLEObject.Application Property (Visio)
keywords: vis_sdr.chm15213090
f1_keywords:
- vis_sdr.chm15213090
ms.prod: visio
api_name:
- Visio.OLEObject.Application
ms.assetid: 402c270d-445d-a067-00cf-5f809eb23e01
ms.date: 06/08/2017
---


# OLEObject.Application Property (Visio)

Returns the instance of Microsoft Visio that is associated with an object. Read-only.


## Syntax

 _expression_ . **Application**

 _expression_ A variable that represents an **OLEObject** object.


### Return Value

Application


## Example

The following Microsoft Visual Basic for Applications (VBA) macro gets the  **Application** object associated with the active document and prints its process ID number in the Immediate window.


```vb
 
Public Sub Application_Example() 
 
 Dim vsoApplication As Visio.Application 
 Dim vsoDocument As Visio.Document 
 
 Set vsoDocument = ActiveDocument 
 
 'Get the instance of Visio associated with the Document object. 
 Set vsoApplication = vsoDocument.Application 
 Debug.Print "The process ID of the Application object associated with the active document is: " &; vsoApplication.ProcessID 
 
End Sub
```


