---
title: Validation.Application Property (Visio)
keywords: vis_sdr.chm18013090
f1_keywords:
- vis_sdr.chm18013090
ms.prod: visio
api_name:
- Visio.Validation.Application
ms.assetid: 42d03033-f09c-09f0-1458-0cf4afa546b3
ms.date: 06/08/2017
---


# Validation.Application Property (Visio)

Returns the instance of Microsoft Visio that is associated with an object. Read-only.


## Syntax

 _expression_ . **Application**

 _expression_ A variable that represents a **[Validation](validation-object-visio.md)** object.


### Return Value

 **[Application](application-object-visio.md)**


## Example

The following Visual Basic for Applications (VBA) macro gets the  **Application** object associated with the **Validation** object for the active document and prints its process ID number in the **Immediate** window.


```vb
Public Sub Application_Example() 
 
 Dim vsoApplication As Visio.Application 
 Dim vsoValidation As Validation 
 
 Set vsoValidation = ActiveDocument.Validation 
 
 'Get the instance of Visio associated with the Validation object. 
 Set vsoApplication = vsoValidation.Application 
 Debug.Print "The process ID of the Application object associated with the active document is: " &; vsoApplication.ProcessID 
 
End Sub
```


