---
title: Workbook.ProtectSharing Method (Excel)
keywords: vbaxl10.chm199265
f1_keywords:
- vbaxl10.chm199265
ms.prod: excel
api_name:
- Excel.Workbook.ProtectSharing
ms.assetid: 26660bc6-136a-ffc8-987e-c96db9c08231
ms.date: 06/08/2017
---


# Workbook.ProtectSharing Method (Excel)

Saves the workbook and protects it for sharing.


## Syntax

 _expression_ . **ProtectSharing**( **_Filename_** , **_Password_** , **_WriteResPassword_** , **_ReadOnlyRecommended_** , **_CreateBackup_** , **_SharingPassword_** , **_FileFormat_** )

 _expression_ A variable that represents a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Filename_|Optional| **Variant**|A string indicating the name of the saved file. You can include a full path; if you don?t, Microsoft Excel saves the file in the current folder.|
| _Password_|Optional| **Variant**|A case-sensitive string indicating the protection password to be given to the file. Should be no longer than 15 characters.|
| _WriteResPassword_|Optional| **Variant**|A string indicating the write-reservation password for this file. If a file is saved with the password and the password isn?t supplied when the file is opened, the file is opened read-only.|
| _ReadOnlyRecommended_|Optional| **Variant**| **True** to display a message when the file is opened, recommending that the file be opened read-only.|
| _CreateBackup_|Optional| **Variant**| **True** to create a backup file.|
| _SharingPassword_|Optional| **Variant**|A string indicating the password to be used to protect the file for sharing.|
| _FileFormat_|Optional| **Variant**|A string indicating the file format.|

## Remarks

Use strong passwords that combine uppercase and lowercase letters, numbers, and symbols. Weak passwords don't mix these elements. Strong password: Y6dh!et5. Weak password: House27. Passwords should be 8 or more characters in length. A pass phrase that uses 14 or more characters is better. For more information, see Help protect your personal information with strong passwords. It is critical that you remember your password. If you forget your password, Microsoft cannot retrieve it. Store the passwords that you write down in a secure place away from the information that they help protect. 


## Example

This example saves workbook one and protects it for sharing.


```vb
 
Sub ProtectWorkbook() 
 
    Dim wbAWB As Workbook 
    Dim strPwd As String 
    Dim strSharePwd As String 
 
    Set wbAWB = Application.ActiveWorkbook 
 
    strPwd = InputBox("Enter password for the file") 
    strSharePwd = InputBox("Enter password for sharing") 
 
    wbAWB.ProtectSharing Password:=strPwd, _ 
        SharingPassword:=strSharePwd 
 
End Sub
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

