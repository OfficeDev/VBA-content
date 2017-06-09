---
title: Workbook.PasswordEncryptionFileProperties Property (Excel)
keywords: vbaxl10.chm199215
f1_keywords:
- vbaxl10.chm199215
ms.prod: excel
api_name:
- Excel.Workbook.PasswordEncryptionFileProperties
ms.assetid: 536ad729-424e-5f81-85e9-8a6ed71fb11a
ms.date: 06/08/2017
---


# Workbook.PasswordEncryptionFileProperties Property (Excel)

 **True** if Microsoft Excel encrypts file properties for the specified password-protected workbook. Read-only **Boolean** .


## Syntax

 _expression_ . **PasswordEncryptionFileProperties**

 _expression_ A variable that represents a **Workbook** object.


## Remarks

Use the  **[SetPasswordEncryptionOptions](workbook-setpasswordencryptionoptions-method-excel.md)** method to specify whether Excel encrypts file properties for the specified password-protected workbook.


## Example

This example sets the password encryption options if the file properties are not encrypted for password-protected workbooks.


```vb
Sub SetPasswordOptions() 
 
 With ActiveWorkbook 
 If .PasswordEncryptionFileProperties = False Then 
 .SetPasswordEncryptionOptions _ 
 PasswordEncryptionProvider:="Microsoft RSA SChannel Cryptographic Provider", _ 
 PasswordEncryptionAlgorithm:="RC4", _ 
 PasswordEncryptionKeyLength:=56, _ 
 PasswordEncryptionFileProperties:=True 
 End If 
 End With 
 
End Sub
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

