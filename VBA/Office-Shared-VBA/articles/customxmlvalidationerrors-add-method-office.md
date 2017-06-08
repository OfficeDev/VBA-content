---
title: CustomXMLValidationErrors.Add Method (Office)
keywords: vbaof11.chm308004
f1_keywords:
- vbaof11.chm308004
ms.prod: office
api_name:
- Office.CustomXMLValidationErrors.Add
ms.assetid: 21b330f2-9c4e-7216-cebb-70d602d68279
ms.date: 06/08/2017
---


# CustomXMLValidationErrors.Add Method (Office)

Adds a  **CustomXMLValidationError** object containing an XML validation error to the **CustomXMLValidationErrors** collection.


## Syntax

 _expression_. **Add**( **_Node_**, **_ErrorName_**, **_ErrorText_**, **_ClearedOnUpdate_** )

 _expression_ An expression that returns a **CustomXMLValidationErrors** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Node_|Required|**CustomXMLNode**|Represents the node where the error occurred.|
| _ErrorName_|Required|**String**|Contains the name of the error.|
| _ErrorText_|Optional|**String**|Contains the descriptive error text.|
| _ClearedOnUpdate_|Optional|**Boolean**|Specifies whether the error is to be cleared from the  **CustomXMLValidationErrors** collection when the XML is corrected and updated.|

## Example

The following example adds an error message to the collection.


```
Dim objCustomXMLValidationErrors as CustomXMLValidationErrors 
 
On Error GoTo Err 
 
' Adds the specified error message to the collection. 
objCustomXMLValidationErrors.Add("//badTag", "ValidationError", "To add content to this stream, it must be valid, well-formed XML.", True) 

```


## See also


#### Concepts


[CustomXMLValidationErrors Object](customxmlvalidationerrors-object-office.md)
#### Other resources


[CustomXMLValidationErrors Object Members](customxmlvalidationerrors-members-office.md)

