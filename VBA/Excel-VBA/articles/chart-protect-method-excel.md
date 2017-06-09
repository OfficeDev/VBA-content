---
title: Chart.Protect Method (Excel)
keywords: vbaxl10.chm149173
f1_keywords:
- vbaxl10.chm149173
ms.prod: excel
api_name:
- Excel.Chart.Protect
ms.assetid: 5f46d721-021b-d615-12c6-78aab49df500
ms.date: 06/08/2017
---


# Chart.Protect Method (Excel)

Protects a chart so that it cannot be modified.


## Syntax

 _expression_ . **Protect**( **_Password_** , **_DrawingObjects_** , **_Contents_** , **_Scenarios_** , **_UserInterfaceOnly_** )

 _expression_ A variable that represents a **Chart** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Password_|Optional| **Variant**|A string that specifies a case-sensitive password for the worksheet or workbook. If this argument is omitted, you can unprotect the worksheet or workbook without using a password. Otherwise, you must specify the password to unprotect the worksheet or workbook. If you forget the password, you cannot unprotect the worksheet or workbook. Use strong passwords that combine uppercase and lowercase letters, numbers, and symbols. Weak passwords don't mix these elements. Strong password: Y6dh!et5. Weak password: House27. Passwords should be 8 or more characters in length. A pass phrase that uses 14 or more characters is better. For more information, see Help protect your personal information with strong passwords. It is critical that you remember your password. If you forget your password, Microsoft cannot retrieve it. Store the passwords that you write down in a secure place away from the information that they help protect. |
| _DrawingObjects_|Optional| **Variant**| **True** to protect shapes. The default value is **True** .|
| _Contents_|Optional| **Variant**| **True** to protect contents. For a chart, this protects the entire chart. For a worksheet, this protects the locked cells. The default value is **True** .|
| _Scenarios_|Optional| **Variant**| **True** to protect scenarios. This argument is valid only for worksheets. The default value is **True** .|
| _UserInterfaceOnly_|Optional| **Variant**| **True** to protect the user interface, but not macros. If this argument is omitted, protection applies both to macros and to the user interface.|

## See also


#### Concepts


[Chart Object](chart-object-excel.md)

