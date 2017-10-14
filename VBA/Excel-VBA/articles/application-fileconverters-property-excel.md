---
title: Application.FileConverters Property (Excel)
keywords: vbaxl10.chm133134
f1_keywords:
- vbaxl10.chm133134
ms.prod: excel
api_name:
- Excel.Application.FileConverters
ms.assetid: 7aebb0b3-6143-8dce-9893-e15decfe1c09
ms.date: 06/08/2017
---


# Application.FileConverters Property (Excel)

Returns information about installed file converters. Returns  **null** if there are no converters installed. Read-only **Variant** .


## Syntax

 _expression_ . **FileConverters**( **_Index1_** , **_Index2_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index1_|Optional| **Variant**|The long name of the converter, including the file-type search string in Windows (for example, "Lotus 1-2-3 Files (*.wk*)").|
| _Index2_|Optional| **Variant**|The path of the converter DLL or code resource.|

## Remarks

If you don?t specify the index arguments, this property returns an array that containing information about all the installed file converters. Each row in the array contains information about a single file converter, as shown in the following table.



|**Column**|**Contents**|
|:-----|:-----|
|1|The long name of the converter|
|2|The path of the converter DLL or code resource|
|3|The file-extension search string|

## Example

This example displays a message if the Multiplan file converter is installed.


```vb
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


[Application Object](application-object-excel.md)

