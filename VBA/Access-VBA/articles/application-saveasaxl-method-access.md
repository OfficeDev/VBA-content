---
title: Application.SaveAsAXL Method (Access)
keywords: vbaac10.chm14664
f1_keywords:
- vbaac10.chm14664
ms.prod: access
api_name:
- Access.Application.SaveAsAXL
ms.assetid: a9557499-7e69-b405-8e2f-d9fcb23fb012
ms.date: 06/08/2017
---


# Application.SaveAsAXL Method (Access)

Exports the specified object to an Application XML (AXL) file.


## Syntax

 _expression_. **SaveAsAXL**( ** _ObjectType_**, ** _ObjectName_**, ** _FileName_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ObjectType_|Required|**AcObjectType**|Specifies the type of object to export.|
| _ObjectName_|Required|**String**|Specifies the name of the object to export. |
| _FileName_|Required|**String**|Specifies the full path and filename of the AXL file to create.|

## Remarks

The  **SaveAsAXL** method does not provide a warning when the file specified in the _FileName_ argument already exists. If this occurs, the file will be overwritten.

The  **SaveAsAXL** method generates a run-time error if the current database is not a Web database.

For more information about AXL, see [[MS-AXL]: Access Application Transfer Protocol Structure Specification](http://msdn.microsoft.com/en-us/library/dd927584.aspx).


## See also


#### Concepts


[Application Object](application-object-access.md)

