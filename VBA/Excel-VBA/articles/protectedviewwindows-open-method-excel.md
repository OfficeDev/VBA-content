---
title: ProtectedViewWindows.Open Method (Excel)
keywords: vbaxl10.chm913077
f1_keywords:
- vbaxl10.chm913077
ms.prod: excel
api_name:
- Excel.ProtectedViewWindows.Open
ms.assetid: bb003d53-949e-842a-f6f1-3ca30f396837
ms.date: 06/08/2017
---


# ProtectedViewWindows.Open Method (Excel)

Opens the specified workbook in a new  **Protected View** window.


## Syntax

 _expression_ . **Open**( **_Filename_** , **_Password_** , **_AddToMru_** , **_RepairMode_** )

 _expression_ A variable that represents a **[ProtectedViewWindows](protectedviewwindows-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Filename_|Required| **String**|The name of the workbook (paths are accepted).|
| _Password_|Optional| **Variant**|The password for opening the workbook.|
| _AddToMru_|Optional| **Variant**| **True** to add the file name to the list of recently used files on the **Recent** tab of the **Backstage** view.|
| _RepairMode_|Optional| **Variant**| **True** to repair the workbook to prevent file corruption.|

### Return Value

 **[ProtectedViewWindow](protectedviewwindow-object-excel.md)**


## Remarks

Avoid using hard-coded passwords in your applications. If a password is required in a procedure, request the password from the user, store it in a variable, and then use the variable in your code.


## Example

The following code example opens a workbook in a new  **Protected View** window.


```vb
ProtectedViewWindows.Open FileName:="C:\MyFiles\MyWorkbook.xls" 

```


## See also


#### Concepts


[ProtectedViewWindows Object](protectedviewwindows-object-excel.md)

