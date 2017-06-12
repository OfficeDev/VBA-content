---
title: References.AddFromGuid Method (Access)
keywords: vbaac10.chm12642
f1_keywords:
- vbaac10.chm12642
ms.prod: access
api_name:
- Access.References.AddFromGuid
ms.assetid: df383ef3-e27c-9590-2ee7-d078060c9313
ms.date: 06/08/2017
---


# References.AddFromGuid Method (Access)

The  **AddFromGUID** method creates a **[Reference](reference-object-access.md)** object based on the GUID that identifies a type library. **Reference** object.


## Syntax

 _expression_. **AddFromGuid**( ** _Guid_**, ** _Major_**, ** _Minor_** )

 _expression_ A variable that represents a **References** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Guid_|Required|**String**|A GUID that identifies a type library.|
| _Major_|Required|**Long**|The major version number of the reference.|
| _Minor_|Required|**Long**|The minor version number of the reference.|

### Return Value

Reference


## Remarks

The  **[GUID](reference-guid-property-access.md)** property returns the GUID for a specified **Reference** object. If you've stored the value of the **GUID** property, you can use it to re-create a reference that's been broken.


## Example

The following example re-creates a reference to the  **Microsoft Scripting Runtime** version 1.0, based on its GUID on the user's system.


```vb
References.AddFromGuid "{420B2830-E718-11CF-893D-00A0C9054228}", 1, 0
```


## See also


#### Concepts


[References Collection](references-object-access.md)

