---
title: Worksheet.Protect Method (Excel)
keywords: vbaxl10.chm175156
f1_keywords:
- vbaxl10.chm175156
ms.prod: excel
api_name:
- Excel.Worksheet.Protect
ms.assetid: ed517a80-eea9-4268-5fbc-69c659beac0e
ms.date: 06/08/2017
---


# Worksheet.Protect Method (Excel)

Protects a worksheet so that it cannot be modified.


## Syntax

 _expression_ . **Protect**( **_Password_** , **_DrawingObjects_** , **_Contents_** , **_Scenarios_** , **_UserInterfaceOnly_** , **_AllowFormattingCells_** , **_AllowFormattingColumns_** , **_AllowFormattingRows_** , **_AllowInsertingColumns_** , **_AllowInsertingRows_** , **_AllowInsertingHyperlinks_** , **_AllowDeletingColumns_** , **_AllowDeletingRows_** , **_AllowSorting_** , **_AllowFiltering_** , **_AllowUsingPivotTables_** )

 _expression_ A variable that represents a **Worksheet** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Password_|Optional| **Variant**|A string that specifies a case-sensitive password for the worksheet or workbook. If this argument is omitted, you can unprotect the worksheet or workbook without using a password. Otherwise, you must specify the password to unprotect the worksheet or workbook. If you forget the password, you cannot unprotect the worksheet or workbook. Use strong passwords that combine uppercase and lowercase letters, numbers, and symbols. Weak passwords don't mix these elements. Strong password: Y6dh!et5. Weak password: House27. Passwords should be 8 or more characters in length. A pass phrase that uses 14 or more characters is better. For more information, see Help protect your personal information with strong passwords. It is critical that you remember your password. If you forget your password, Microsoft cannot retrieve it. Store the passwords that you write down in a secure place away from the information that they help protect. |
| _DrawingObjects_|Optional| **Variant**| **True** to protect shapes. The default value is **True** .|
| _Contents_|Optional| **Variant**| **True** to protect contents. For a chart, this protects the entire chart. For a worksheet, this protects the locked cells. The default value is **True** .|
| _Scenarios_|Optional| **Variant**| **True** to protect scenarios. This argument is valid only for worksheets. The default value is **True** .|
| _UserInterfaceOnly_|Optional| **Variant**| **True** to protect the user interface, but not macros. If this argument is omitted, protection applies both to macros and to the user interface.|
| _AllowFormattingCells_|Optional| **Variant**| **True** allows the user to format any cell on a protected worksheet. The default value is **False** .|
| _AllowFormattingColumns_|Optional| **Variant**| **True** allows the user to format any column on a protected worksheet. The default value is **False** .|
| _AllowFormattingRows_|Optional| **Variant**| **True** allows the user to format any row on a protected. The default value is **False** .|
| _AllowInsertingColumns_|Optional| **Variant**| **True** allows the user to insert columns on the protected worksheet. The default value is **False** .|
| _AllowInsertingRows_|Optional| **Variant**| **True** allows the user to insert rows on the protected worksheet. The default value is **False** .|
| _AllowInsertingHyperlinks_|Optional| **Variant**| **True** allows the user to insert hyperlinks on the worksheet. The default value is **False** .|
| _AllowDeletingColumns_|Optional| **Variant**| **True** allows the user to delete columns on the protected worksheet, where every cell in the column to be deleted is unlocked. The default value is **False** .|
| _AllowDeletingRows_|Optional| **Variant**| **True** allows the user to delete rows on the protected worksheet, where every cell in the row to be deleted is unlocked. The default value is **False** .|
| _AllowSorting_|Optional| **Variant**| **True** allows the user to sort on the protected worksheet. Every cell in the sort range must be unlocked or unprotected. The default value is **False** .|
| _AllowFiltering_|Optional| **Variant**| **True** allows the user to set filters on the protected worksheet. Users can change filter criteria but can not enable or disable an auto filter. Users can set filters on an existing auto filter. The default value is **False** .|
| _AllowUsingPivotTables_|Optional| **Variant**| **True** allows the user to use pivot table reports on the protected worksheet. The default value is **False** .|

## Remarks

If changes wanted to be made to a protected worksheet, it is possible to use the  **Protect** method on a protected worksheet if the password is supplied. Also, another method would be to unprotect the worksheet, make the necessary changes, and then protect the worksheet again.


 **Note**  'Unprotected' means the cell may be locked ( **Format Cells** dialog box) but is included in a range defined in the **Allow Users to Edit Ranges** dialog box, and the user has unprotected the range with a password or been validated via NT permissions.


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

