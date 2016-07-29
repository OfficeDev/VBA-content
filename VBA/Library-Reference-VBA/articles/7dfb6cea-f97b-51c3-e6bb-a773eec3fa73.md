
# EncryptionProvider.Save Method (Office)

Saves an encrypted document.


## Syntax

 _expression_. **Save**( ** _SessionHandle_**, ** _EncryptionData_** )

 _expression_ An expression that returns a **EncryptionProvider** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SessionHandle_|Required|**Long**|The ID of the current session.|
| _EncryptionData_|Required|**IUnknown**|Contains the encryption information.|

### Return Value

Long


## Remarks

When you save a file to the Office Open XML File Format (which is the only format that supports custom file encryption), then the provider is called by your COM add-in to encrypt the document. If you attempt to save to a format that does not support custom file encryption and you have the appropriate rights to do so, then Microsoft Office will save the document without encryption. This allows documents to be exported to formats that do not support encryption or rights management.


## See also


#### Concepts


[EncryptionProvider Object](9f5cc550-6bcb-2748-14a7-696cf8ef021b.md)
#### Other resources


[EncryptionProvider Object Members](48bed5b8-b284-4b52-4143-153ae1c751a4.md)