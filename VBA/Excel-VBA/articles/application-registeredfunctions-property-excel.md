---
title: Application.RegisteredFunctions Property (Excel)
keywords: vbaxl10.chm133198
f1_keywords:
- vbaxl10.chm133198
ms.prod: excel
api_name:
- Excel.Application.RegisteredFunctions
ms.assetid: c8922122-7de8-ebbb-0dfd-1dfe3974278e
ms.date: 06/08/2017
---


# Application.RegisteredFunctions Property (Excel)

Returns information about functions in either dynamic-link libraries (DLLs) or code resources that were registered with the REGISTER or REGISTER.ID macro functions. Read-only  **Variant** .


## Syntax

 _expression_ . **RegisteredFunctions**( **_Index1_** , **_Index2_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index1_|Optional| **Variant**|The name of the DLL or code resource.|
| _Index2_|Optional| **Variant**|The name of the function.|

## Remarks

If you don?t specify the index arguments, this property returns an array that contains a list of all registered functions. Each row in the array contains information about a single function, as shown in the following table.



|**Column**|**Contents**|
|:-----|:-----|
|1|The name of the DLL or code resource|
|2|The name of the procedure in the DLL or code resource|
|3|Strings specifying the data types of the return values, and the number and data types of the arguments|
If there are no registered functions, this property returns  **null** .


## Example

This example creates a list of registered functions, placing one registered function in each row on Sheet1. Column A contains the full path and file name of the DLL or code resource, column B contains the function name, and column C contains the argument data type code.


```vb
theArray = Application.RegisteredFunctions 
If IsNull(theArray) Then 
 MsgBox "No registered functions" 
Else 
 For i = LBound(theArray) To UBound(theArray) 
 For j = 1 To 3 
 Worksheets("Sheet1").Cells(i, j). _ 
 Formula = theArray(i, j) 
 Next j 
 Next i 
End If
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

