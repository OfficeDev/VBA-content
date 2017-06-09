---
title: PropertyAccessor.DeleteProperty Method (Outlook)
keywords: vbaol11.chm1978
f1_keywords:
- vbaol11.chm1978
ms.prod: outlook
api_name:
- Outlook.PropertyAccessor.DeleteProperty
ms.assetid: 9acb52b5-13a7-7363-7e17-83804037f33b
ms.date: 06/08/2017
---


# PropertyAccessor.DeleteProperty Method (Outlook)

Deletes the property specified by  _SchemaName_ .


## Syntax

 _expression_ . **DeleteProperty**( **_SchemaName_** )

 _expression_ A variable that represents a **PropertyAccessor** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SchemaName_|Required| **String**|The name of the property that is to be deleted for the parent object of the  **[PropertyAccessor](propertyaccessor-object-outlook.md)** object. The property is referenced by namespace. For more information, see[Referencing Properties by Namespace](http://msdn.microsoft.com/library/c1c7bfa9-64d7-81d2-84e7-f0a4c57780b3%28Office.15%29.aspx).|

## Remarks

The caller must have the permission to delete properties. The  **DeleteProperty** method deletes only custom properties; it does not delete any Outlook built-in property or any MAPI property. It does not delete custom properties of the **[DocumentItem](documentitem-object-outlook.md)** object.


## See also


#### Concepts


[PropertyAccessor Object](propertyaccessor-object-outlook.md)

