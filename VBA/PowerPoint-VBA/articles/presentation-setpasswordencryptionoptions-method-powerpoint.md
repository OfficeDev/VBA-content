---
title: Presentation.SetPasswordEncryptionOptions Method (PowerPoint)
keywords: vbapp10.chm583079
f1_keywords:
- vbapp10.chm583079
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.SetPasswordEncryptionOptions
ms.assetid: 03c07952-784b-eba6-af71-57d3d1414f81
ms.date: 06/08/2017
---


# Presentation.SetPasswordEncryptionOptions Method (PowerPoint)

Sets the options Microsoft PowerPoint uses for encrypting presentations with passwords.


## Syntax

 _expression_. **SetPasswordEncryptionOptions**( **_PasswordEncryptionProvider_**, **_PasswordEncryptionAlgorithm_**, **_PasswordEncryptionKeyLength_**, **_PasswordEncryptionFileProperties_** )

 _expression_ A variable that represents a **Presentation** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _PasswordEncryptionProvider_|Required|**String**|The name of the encryption provider.|
| _PasswordEncryptionAlgorithm_|Required|**String**|The name of the encryption algorithm. PowerPoint supports stream-encrypted algorithms.|
| _PasswordEncryptionKeyLength_|Required|**Long**|The encryption key length. Must be a multiple of 8, starting at 40.|
| _PasswordEncryptionFileProperties_|Required|**MsoTriState**|**msoTrue** for PowerPoint to encrypt file properties.|

## Remarks

The  _PasswordEncryptionFileProperties_ parameter value can be one of these **MsoTriState** constants.


||
|:-----|
|**msoFalse**|
|**msoTrue**|

## Example

This example sets the password encryption options if the file properties are not encrypted for password-protected documents.


```vb
Sub PasswordSettings()

    With ActivePresentation
        If .PasswordEncryptionFileProperties = msoFalse Then
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


[Presentation Object](presentation-object-powerpoint.md)

