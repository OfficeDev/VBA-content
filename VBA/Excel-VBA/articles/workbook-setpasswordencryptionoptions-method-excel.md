---
title: Workbook.SetPasswordEncryptionOptions Method (Excel)
keywords: vbaxl10.chm199214
f1_keywords:
- vbaxl10.chm199214
ms.prod: excel
api_name:
- Excel.Workbook.SetPasswordEncryptionOptions
ms.assetid: 3b6c9bfe-4cfb-1dde-fd57-07dd474df7db
ms.date: 06/08/2017
---


# Workbook.SetPasswordEncryptionOptions Method (Excel)

Sets the options for encrypting workbooks using passwords.


## Syntax

 _expression_ . **SetPasswordEncryptionOptions**( **_PasswordEncryptionProvider_** , **_PasswordEncryptionAlgorithm_** , **_PasswordEncryptionKeyLength_** , **_PasswordEncryptionFileProperties_** )

 _expression_ A variable that represents a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _PasswordEncryptionProvider_|Optional| **Variant**|A case sensitive string of the encryption provider.|
| _PasswordEncryptionAlgorithm_|Optional| **Variant**|A case sensitive string of the algorithmic short name (i.e. "RC4").|
| _PasswordEncryptionKeyLength_|Optional| **Variant**|The encryption key length which is a multiple of 8 (40 or greater).|
| _PasswordEncryptionFileProperties_|Optional| **Variant**| **True** (default) to encrypt file properties.|

## Remarks

The  _PasswordEncryptionProvider_,  _PasswordEncryptionAlgorithm_, and  _PasswordEncryptionKeyLength_ arguments are not independent of each other. A selected encryption provider limits the set of algorithms and key length that can be chosen.

For the  _PasswordEncryptionKeyLength_ argument there is no inherent limit on the range of the key length. The range is determined by the Cryptographic Service Provider which also determines the cryptographic algorithm.


## Example

This example sets the password encryption options for the active workbook.


```vb
Sub SetPasswordOptions() 
 
 ActiveWorkbook.SetPasswordEncryptionOptions _ 
 PasswordEncryptionProvider:="Microsoft RSA SChannel Cryptographic Provider", _ 
 PasswordEncryptionAlgorithm:="RC4", _ 
 PasswordEncryptionKeyLength:=56, _ 
 PasswordEncryptionFileProperties:=True 
 
End Sub
```


 **Note**  The code and this method do not do anything for the new Excel file formats (xlsx, xlsb, xlsm, etc...) as the workbook will always use AES 128 bit encryption. If a property is set using this method, it will appear set. When the file is reloaded, the properties will be reset to the AES setting.


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

