---
title: Application.FileConverters Property (PowerPoint)
keywords: vbapp10.chm502060
f1_keywords:
- vbapp10.chm502060
ms.prod: powerpoint
api_name:
- PowerPoint.Application.FileConverters
ms.assetid: 2eaa06eb-e32c-cf07-03a2-880048468188
ms.date: 06/08/2017
---


# Application.FileConverters Property (PowerPoint)

Returns information about installed file converters. Returns  **null** if there are no converters installed. Read-only **Variant**.


## Syntax

 _expression_. **FileConverters**( **_Index1_**, **_Index2_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index1_|Optional|**Variant**|The long name of the converter, including the file-type search string in Windows (for example, "Lotus 1-2-3 Files (*.wk*)").|
| _Index2_|Optional|**Variant**|The path of the converter DLL or code resource.|

## Remarks

If you do not specify the index arguments, this property returns an array that contains information about all the installed file converters. Each row in the array contains information about a single file converter, as shown in the following table.



|**Column**|**Contents**|
|:-----|:-----|
|1|The long name of the converter|
|2|The path of the converter DLL or code resource|
|3|The file-extension search string|

## Example

The following example displays a message if the Multiplan file converter is installed.


```
installedCvts = Application.FileConverters

foundMultiplan = False

If Not IsNull(installedCvts) Then

    For arrayRow = 1 To UBound(installedCvts, 1)

        If installedCvts(arrayRow, 1) Like "*Multiplan*" Then

            foundMultiplan = True

            Exit For

        End If

    Next arrayRow

End If

If foundMultiplan = True Then

    MsgBox "Multiplan converter is installed"

Else

    MsgBox "Multiplan converter is not installed"

End If
```


## See also


#### Concepts


[Application Object](application-object-powerpoint.md)

