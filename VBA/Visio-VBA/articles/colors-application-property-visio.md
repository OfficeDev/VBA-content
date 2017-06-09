---
title: Colors.Application Property (Visio)
keywords: vis_sdr.chm12313090
f1_keywords:
- vis_sdr.chm12313090
ms.prod: visio
api_name:
- Visio.Colors.Application
ms.assetid: 89418804-bf4b-d322-e0e1-84c8817b419a
ms.date: 06/08/2017
---


# Colors.Application Property (Visio)

Returns the instance of Microsoft Visio that is associated with an object. Read-only.


## Syntax

 _expression_ . **Application**

 _expression_ A variable that represents a **Colors** object.


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


