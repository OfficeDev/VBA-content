
# Workbook.SetPasswordEncryptionOptions Method (Excel)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Sets the options for encrypting workbooks using passwords.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **SetPasswordEncryptionOptions**( **_PasswordEncryptionProvider_**,  **_PasswordEncryptionAlgorithm_**,  **_PasswordEncryptionKeyLength_**,  **_PasswordEncryptionFileProperties_**)

 _expression_A variable that represents a  **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|PasswordEncryptionProvider|Optional| **Variant**|A case sensitive string of the encryption provider.|
|PasswordEncryptionAlgorithm|Optional| **Variant**|A case sensitive string of the algorithmic short name (i.e. "RC4").|
|PasswordEncryptionKeyLength|Optional| **Variant**|The encryption key length which is a multiple of 8 (40 or greater).|
|PasswordEncryptionFileProperties|Optional| **Variant**| **True** (default) to encrypt file properties.|

## Remarks
<a name="sectionSection1"> </a>

The PasswordEncryptionProvider, PasswordEncryptionAlgorithm, and PasswordEncryptionKeyLength arguments are not independent of each other. A selected encryption provider limits the set of algorithms and key length that can be chosen.

For the PasswordEncryptionKeyLength argument there is no inherent limit on the range of the key length. The range is determined by the Cryptographic Service Provider which also determines the cryptographic algorithm.


## Example
<a name="sectionSection2"> </a>

This example sets the password encryption options for the active workbook.


```
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
<a name="sectionSection2"> </a>


#### Concepts


 [Workbook Object](8c00aa60-c974-eed3-0812-3c9625eb0d4c.md)
#### Other resources


 [Workbook Object Members](dce102a3-25de-3ff4-2ce5-bc56e08baca7.md)
