---
title: Application.Version Property (Visio)
keywords: vis_sdr.chm10014640
f1_keywords:
- vis_sdr.chm10014640
ms.prod: visio
api_name:
- Visio.Application.Version
ms.assetid: c2e3b022-507d-c73c-6fa4-9689cc5600f3
ms.date: 06/08/2017
---


# Application.Version Property (Visio)

Returns the version of a running Microsoft Visio instance. Read-only.


## Syntax

 _expression_ . **Version**

 _expression_ A variable that represents an **Application** object.


### Return Value

String


## Remarks

Use the  **Version** property of the **Application** object to verify the version of a particular Visio instance. This information is helpful if your program requires a particular version. Both the major and minor version numbers are returned. The string returned by Visio is 15.0.


## Example

This Microsoft Visual Basic for Applications (VBA) program shows how to print the version of a Visio instance in the Immediate window.


```vb
 
Public Sub Version_Example() 
 
 Dim vsoApplication As Visio.Application 
 Dim strVersion As String 
 Dim intDotPosition As Integer 
 Set vsoApplication = CreateObject("Visio.Application") 
 
 strVersion = vsoApplication.Version 
 intDotPosition = InStr(strVersion, ".") 
 
 Debug.Print " Major Version : "; Left(strVersion, intDotPosition - 1) 
 Debug.Print " Minor Version : "; Right(strVersion, Len(strVersion) - intDotPosition) 
 
End Sub
```


