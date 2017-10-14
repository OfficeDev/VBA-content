---
title: InvisibleApp.FullBuild Property (Visio)
keywords: vis_sdr.chm17551220
f1_keywords:
- vis_sdr.chm17551220
ms.prod: visio
api_name:
- Visio.InvisibleApp.FullBuild
ms.assetid: f5c7fc4a-6627-888b-0465-0e251936b7b6
ms.date: 06/08/2017
---


# InvisibleApp.FullBuild Property (Visio)

Returns the full build number of the running instance. Read-only.


## Syntax

 _expression_ . **FullBuild**

 _expression_ A variable that represents an **InvisibleApp** object.


### Return Value

Long


## Remarks

The format of the build number is described in the following table.



|** Bits**|** Description**|
|:-----|:-----|
|0 - 15|Internal build number|
|16 - 20|Internal revision number|
|21 - 25|Minor version number|
|26 - 30|Major version number (Visio = 15)|
|31|Reserved|
In addition, for Visio, to obtain the correct full build number, it is necessary to add 1000 to the internal revision number part of the full build number returned by the  **FullBuild** property, as shown in the following macro.

The build number of the running instance is written to the  **FullBuildNumberCreated** property when a new document is created, and to the **FullBuildNumberEdited** property when a document is edited.


## Example

The following Microsoft Visual Basic procedures show how to use the  **FullBuild** property to get the full build number of the current instance of Visio. Once the full build number has been obtained, the **ParseFullBuildProperty** procedure parses the number and prints the results in the **Immediate** window.


```vb
Public Sub FullBuild_Example() 
 
 Dim lngFullBuild as Long 
 lngFullBuild = Application.FullBuild 
 ParseFullBuildProperty (lngFullBuild) 
 
End Sub
```


```vb
Public Sub ParseFullBuildProperty(ByRef lngFullBuild As Long) 
 
 Dim lngMajor As Long 
 Dim lngMinor As Long 
 Dim lngRevision As Long 
 Dim lngBuild As Long 
 Dim lngNumber As Long 
 
 lngNumber = lngFullBuild 
 
 ' Low 16 bits: 
 lngBuild = lngNumber Mod 65536 
 lngNumber = lngNumber / 65536 
 
 'Next 5 bits: 
 lngRevision = lngNumber Mod 32 
 lngNumber = lngNumber / 32 
 
 'Next 5 bits: 
 lngMinor = lngNumber Mod 32 
 lngNumber = lngNumber / 32 
 
 'Next 5 bits: 
 lngMajor = lngNumber Mod 32 
 lngNumber = lngNumber / 32 
 
 'Remaining 1 bit unused and 0 as of Visio 2010 
 Debug.Print "lngFullBuild (full version specification): " &; lngMajor &; "." _ 
 &; lngMinor &; "." &; lngBuild &; "." &; lngRevision + 1000 
 Debug.Assert(0 = lngNumber) 
 
End Sub
```


