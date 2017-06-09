---
title: PickerResults.Add Method (Office)
keywords: vbaof11.chm339003
f1_keywords:
- vbaof11.chm339003
ms.prod: office
api_name:
- Office.PickerResults.Add
ms.assetid: cf6e4f0f-4373-3caa-ddb3-512ca5c4675f
ms.date: 06/08/2017
---


# PickerResults.Add Method (Office)

Adds a  **PickerResult** object to the **PickerResults** collection.


## Syntax

 _expression_. **Add**( **_Id_**, **_DisplayName_**, **_Type_**, **_SIPId_**, **_ItemData_**, **_SubItems_** )

 _expression_ An expression that returns a **PickerResults** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Id_|Required|**String**|Represents an identifier of the PickerResult.|
| _DisplayName_|Required|**String**|Represents a display name of the PickerResult. |
| _Type_|Required|**String**|Represents a type of the PickerResult.|
| _SIPId_|Optional|**String**|Currently not supported. The SIPId is the identifier for Office Communication Server. It is used only for the people picking scenario.|
| _ItemData_|Optional|**Variant**|Non- displaying item binding data|
| _SubItems_|Optional|**Variant**|Displays the purpose or non-display purpose field data of the PickerResult. It is used for passing column values in the Picker Dialog.|

### Return Value

PickerResult


## Example

The following code sets the Picker Dialog properties and then displays the Picker Dialog.


```
Dim objPickerDialog As PickerDialog 
Dim objPickerProperties As PickerProperties 
Dim objPickerProperty As PickerProperty 
Dim objPickerExistingResults As PickerResults 
Dim objPickerExistingResults As PickerResult 
Dim objPickerResults As PickerResults 
 
' Configure the Picker Dialog properties. 
Set objPickerDialog = Application.PickerDialog 
objPickerDialog.DataHandlerId = "{000CDF0A-0000-0000-C000-000000000046}" 
objPickerDialog.Title = "Sample Picker Dialog" 
Set objPickerProperties = objPickerDialog.Properties 
Set objPickerProperty = objPickerProperties.Add("SiteUrl", "http://my", msoPickerFieldtypeText) 
Set objPickerExistingResults = objPickerDialog.CreatePickerResults 
Set objPickerExistingResult = objPickerExistingResults.Add("johndoe@contoso.com", "John Doe", "User") 
 
' Show the Picker Dialog and get the results. 
Set objPickerResults = objPickerDialog.Show(True, objPickerExistingResult)
```


## See also


#### Concepts


[PickerResults Object](pickerresults-object-office.md)
#### Other resources


[PickerResults Object Members](pickerresults-members-office.md)

