---
title: ProtectedViewWindow.Edit Method (Excel)
keywords: vbaxl10.chm914087
f1_keywords:
- vbaxl10.chm914087
ms.prod: excel
api_name:
- Excel.ProtectedViewWindow.Edit
ms.assetid: bdb626b2-ed4a-06d2-076c-5d242d23a162
ms.date: 06/08/2017
---


# ProtectedViewWindow.Edit Method (Excel)

Opens the workbook that is open in the specified  **Protected View** window for editing.


## Syntax

 _expression_ . **Edit**( **_WriteResPassword_** , **_UpdateLinks_** )

 _expression_ A variable that represents a **[ProtectedViewWindow](protectedviewwindow-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _WriteResPassword_|Optional| **Variant**|The password required to write to a write-reserved workbook. If this argument is omitted and the workbook requires a password, the user will be prompted for the password.|
| _UpdateLinks_|Optional| **Variant**|Specifies the way external references (links) in the file, such as the reference to a range in the Budget.xls workbook in the following formula =SUM([Budget.xls]Annual!C10:C25), are updated. If this argument is omitted, the user is prompted to specify how links will be updated. For more information about the values used by this parameter, see the Remarks section. If Excel is opening a file in the WKS, WK1, or WK3 format and the  _UpdateLinks_ argument is 0, no charts are created; otherwise Excel generates charts from the graphs attached to the file.|

### Return Value

[Workbook](workbook-object-excel.md)


## Remarks

Avoid using hard-coded passwords in your applications. If a password is required in a procedure, request the password from the user, store it in a variable, and then use the variable in your code.

You can specify one of the values, listed in the following table, in the  _UpdateLinks_ parameter to determine whether external references (links) are updated when the workbook is opened.



|**Value**|**Meaning**|
|:-----|:-----|
|0|External references (links) will not be updated when the workbook is opened.|
|3|External references (links) will be updated when the workbook is opened.|

## Example

The following code example opens the workbook that is open in the active  **Protected View** window for editing.


```vb
Dim pvWbk As Workbook 
 
Set pvWbk = ActiveProtectedViewWindow.Edit 

```


## See also


#### Concepts


[ProtectedViewWindow Object](protectedviewwindow-object-excel.md)

