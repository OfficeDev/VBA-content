---
title: CustomXMLPrefixMappings.LookupNamespace Method (Office)
keywords: vbaof11.chm290005
f1_keywords:
- vbaof11.chm290005
ms.prod: office
api_name:
- Office.CustomXMLPrefixMappings.LookupNamespace
ms.assetid: 33a8f054-0e67-0c9e-ce4b-c9d3360df1a6
ms.date: 06/08/2017
---


# CustomXMLPrefixMappings.LookupNamespace Method (Office)

Allows you to get the namespace corresponding to the specified prefix.


## Syntax

 _expression_. **LookupNamespace**( **_Prefix_** )

 _expression_ An expression that returns a **CustomXMLPrefixMappings** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Prefix_|Required|**String**|Contains a prefix in the prefix mapping list.|

### Return Value

String


## Remarks

If no namespace is assigned to the requested prefix, the method returns an empty string ("").


## Example

The following example retrieves the namespace corresponding to the prefix argument.


```
Dim objCustomPrefixMappings As  CustomPrefixMappings 
    Dim strNamespace As String 
 
    ' Gets the namespace corresponding to the prefix. 
   strNamespace = objCustomPrefixMappings.LookupNamespace("xs")
```


## See also


#### Concepts


[CustomXMLPrefixMappings Object](customxmlprefixmappings-object-office.md)
#### Other resources


[CustomXMLPrefixMappings Object Members](customxmlprefixmappings-members-office.md)

