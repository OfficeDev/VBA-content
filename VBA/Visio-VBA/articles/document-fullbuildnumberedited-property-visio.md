---
title: Document.FullBuildNumberEdited Property (Visio)
keywords: vis_sdr.chm10551230
f1_keywords:
- vis_sdr.chm10551230
ms.prod: visio
api_name:
- Visio.Document.FullBuildNumberEdited
ms.assetid: 43a6ff61-2ab8-4e89-0e06-bd2ba6ec0f02
ms.date: 06/08/2017
---


# Document.FullBuildNumberEdited Property (Visio)

Returns the full build number of the instance last used to edit the document. Read-only. 


## Syntax

 _expression_ . **FullBuildNumberEdited**

 _expression_ A variable that represents a **Document** object.


### Return Value

Long


## Remarks

The format of the full build number is described in the following table.



|** Bits**|** Description**|
|:-----|:-----|
| 0 - 15| Internal build number|
| 16 - 20| Internal revision number|
| 21 - 25| Minor version number.|
| 26 - 30| Major version number (Visio = 15)|
| 31| Reserved|
In addition, for Visio, to obtain the correct full build number, it is necessary to add 1000 to the internal revision number part of the full build number returned by the  **FullBuildNumberEdited** property, as shown in the following macro.


## Example

The following Microsoft Visual Basic procedures show how to use the  **FullBuildNumberEdited** property to get the full build number of the instance of Visio last used to edit the document. When the full build number has been obtained, the **ParseFullBuildNumberEditedProperty** procedure parses the number and prints the result in the **Immediate** window.


```vb
Public Sub FullBuildNumberEdited_Example() 
 
 Dim lngFullBuild As Long 
 lngFullBuild = ActiveDocument.FullBuildNumberEdited 
 ParseFullBuildNumberEditedProperty (lngFullBuild) 
 
End Sub 
 
Public Sub ParseFullBuildNumberEditedProperty(ByRef lngFullBuild As Long) 
 
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


