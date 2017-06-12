---
title: EventList.Application Property (Visio)
keywords: vis_sdr.chm12713090
f1_keywords:
- vis_sdr.chm12713090
ms.prod: visio
api_name:
- Visio.EventList.Application
ms.assetid: e63a6499-c7b9-645d-6c7d-dd8d17c046a1
ms.date: 06/08/2017
---


# EventList.Application Property (Visio)

Returns the instance of Microsoft Visio that is associated with an object. Read-only.


## Syntax

 _expression_ . **Application**

 _expression_ A variable that represents an **EventList** object.


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


