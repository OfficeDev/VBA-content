---
title: CustomXMLPrefixMappings.AddNamespace Method (Office)
keywords: vbaof11.chm290004
f1_keywords:
- vbaof11.chm290004
ms.prod: office
api_name:
- Office.CustomXMLPrefixMappings.AddNamespace
ms.assetid: a4a58a81-3fdc-f808-ac19-0eb27e944f29
ms.date: 06/08/2017
---


# CustomXMLPrefixMappings.AddNamespace Method (Office)

Allows you to add a custom namespace/prefix mapping to use when querying an item.


## Syntax

 _expression_. **AddNamespace**( **_Prefix_**, **_NamespaceURI_** )

 _expression_ An expression that returns a **CustomXMLPrefixMappings** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Prefix_|Required|**String**|Contains the prefix to add to the prefix mapping list.|
| _NamespaceURI_|Required|**String**|Contains the namespace to assign to the newly added prefix.|

## Remarks

If the prefix already exists in the Namespace Manager, this method will overwrite the meaning of that prefix except when the prefix is one added or used by the data store ( **IXMLDataStore** interface) internally, in which case it will return an error.


## Example

The following example adds a prefix and namespace to a  **CustomPrefixMappings** object.


```
Sub AddNamespacePrefix() 
  
    Dim objCustomPrefixMappings As  CustomPrefixMappings 
    Dim varCustomMapping As Variant 
 
    ' Adds a custom namespace. 
    varCustomMapping = objCustomPrefixMappings.AddNamespace("xs", "urn:invoice:namespace")      
 
End Sub
```


## See also


#### Concepts


[CustomXMLPrefixMappings Object](customxmlprefixmappings-object-office.md)
#### Other resources


[CustomXMLPrefixMappings Object Members](customxmlprefixmappings-members-office.md)

