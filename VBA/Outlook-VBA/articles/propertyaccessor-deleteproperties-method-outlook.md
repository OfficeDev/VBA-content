---
title: PropertyAccessor.DeleteProperties Method (Outlook)
keywords: vbaol11.chm1979
f1_keywords:
- vbaol11.chm1979
ms.prod: outlook
api_name:
- Outlook.PropertyAccessor.DeleteProperties
ms.assetid: e9c11799-cb75-fd8c-0c98-aca46796bb46
ms.date: 06/08/2017
---


# PropertyAccessor.DeleteProperties Method (Outlook)

Deletes the properties specified in the array  _SchemaNames_ .


## Syntax

 _expression_ . **DeleteProperties**( **_SchemaNames_** )

 _expression_ A variable that represents a **PropertyAccessor** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SchemaNames_|Required| **Variant**|An array that contains the names of the properties that are to be deleted for the parent object of the  **[PropertyAccessor](propertyaccessor-object-outlook.md)** object. These properties are referenced by namespace. For more information, see[Referencing Properties by Namespace](http://msdn.microsoft.com/library/c1c7bfa9-64d7-81d2-84e7-f0a4c57780b3%28Office.15%29.aspx).|

### Return Value

A Variant that is  **Null** ( **Nothing** in VBA) if the operation is successful, or is an array of **[Err](http://msdn.microsoft.com/library/23c9697a-9c6b-18f8-2b86-a0735f082c67%28Office.15%29.aspx)** objects if an error occurs. If the return value is an array, the size of this array is the same as that of the _SchemaNames_ array. An **Err** value in the array is mapped to the error result of deleting the corresponding property in the _SchemaNames_ parameter.


## Remarks

The caller must have the permission to delete properties. The  **DeleteProperties** method only deletes custom properties that exist. It does not delete any Outlook built-in property or any MAPI property. It does not delete custom properties of the **[DocumentItem](documentitem-object-outlook.md)** object.


## See also


#### Concepts


[PropertyAccessor Object](propertyaccessor-object-outlook.md)

