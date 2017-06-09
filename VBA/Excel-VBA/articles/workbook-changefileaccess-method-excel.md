---
title: Workbook.ChangeFileAccess Method (Excel)
keywords: vbaxl10.chm199082
f1_keywords:
- vbaxl10.chm199082
ms.prod: excel
api_name:
- Excel.Workbook.ChangeFileAccess
ms.assetid: 07f9cfc3-eece-efc1-6c03-38782ad7bcc2
ms.date: 06/08/2017
---


# Workbook.ChangeFileAccess Method (Excel)

Changes the access permissions for the workbook. This may require an updated version to be loaded from the disk.


## Syntax

 _expression_ . **ChangeFileAccess**( **_Mode_** , **_WritePassword_** , **_Notify_** )

 _expression_ A variable that represents a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Mode_|Required| **[XlFileAccess](xlfileaccess-enumeration-excel.md)**|Specifies the new access mode.|
| _WritePassword_|Optional| **Variant**|Specifies the write-reserved password if the file is write reserved and  _Mode_ is **xlReadWrite** . Ignored if there's no password for the file or if _Mode_ is **xlReadOnly** .|
| _Notify_|Optional| **Variant**| **True** (or omitted) to notify the user if the file cannot be immediately accessed.|

## Remarks

If you have a file open in read-only mode, you don't have exclusive access to the file. If you change a file from read-only to read/write, Microsoft Excel must load a new copy of the file to ensure that no changes were made while you had the file open as read-only.


## Example

This example sets the active workbook to read-only.


```vb
ActiveWorkbook.ChangeFileAccess Mode:=xlReadOnly
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

