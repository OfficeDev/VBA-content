---
title: Document.FullBuildNumberCreated Property (Visio)
keywords: vis_sdr.chm10551225
f1_keywords:
- vis_sdr.chm10551225
ms.prod: visio
api_name:
- Visio.Document.FullBuildNumberCreated
ms.assetid: 3520525a-4c76-3583-49a6-015f2fb90366
ms.date: 06/08/2017
---


# Document.FullBuildNumberCreated Property (Visio)

Returns the full build number of the instance used to create the document. Read-only.


## Syntax

 _expression_ . **FullBuildNumberCreated**

 _expression_ A variable that represents a **Document** object.


### Return Value

Long


## Remarks

The format of the full build number is described in the following table.



|** Bits**|** Description**|
|:-----|:-----|
| 0 - 15| Internal build number|
| 16 - 20| Internal revision number|
| 21 - 25| Minor version number|
| 26 - 30| Major version number (Visio = 15)|
| 31| Reserved|
In addition, for Visio, to obtain the correct full build number, it is necessary to add 1000 to the internal revision number part of the full build number returned by the  **FullBuildNumberCreated** property, as shown in the following macro.


## Example

The following Microsoft Visual Basic for Applications (VBA) procedures show how to use the  **FullBuildNumberCreated** property to get the full build number of the instance of Visio used to create the document. Once the full build number has been obtained, the **ParseFullBuildNumberCreatedProperty** procedure parses the number and prints the results in the **Immediate** window.


```vb
Public Sub FullBuildNumberCreated_Example() 
 
 Dim lngFullBuild As Long 
 lngFullBuild = ActiveDocument.FullBuildNumberCreated 
 ParseFullBuildNumberCreatedProperty (lngFullBuild) 
 
End Sub 
 
Public Sub ParseFullBuildNumberCreatedProperty(ByRef lngFullBuild As Long) 
 
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


