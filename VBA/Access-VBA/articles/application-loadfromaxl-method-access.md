---
title: Application.LoadFromAXL Method (Access)
keywords: vbaac10.chm14665
f1_keywords:
- vbaac10.chm14665
ms.prod: access
api_name:
- Access.Application.LoadFromAXL
ms.assetid: 1cce0568-1966-c089-a741-b0934b8676d6
ms.date: 06/08/2017
---


# Application.LoadFromAXL Method (Access)

Imports the object defined in an Application XML (AXL) file into the database. 


## Syntax

 _expression_. **LoadFromAXL**( ** _ObjectType_**, ** _ObjectName_**, ** _FileName_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ObjectType_|Required|**AcObjectType**|Specifies the type of object to create.|
| _ObjectName_|Required|**String**|Specifies the name of the object.|
| _FileName_|Required|**String**|Specifies the full path and filename of the AXL file to import.|

## Remarks

The  **LoadFromAXL** method does not provide a warning when the object specified in the _ObjectName_ argument already exists. If an object of the same name already exists, it will be replaced by the object specified in the _ObjectName_ argument.

For more information about AXL, see [[MS-AXL]: Access Application Transfer Protocol Structure Specification](http://msdn.microsoft.com/en-us/library/dd927584.aspx).


## See also


#### Concepts


[Application Object](application-object-access.md)

