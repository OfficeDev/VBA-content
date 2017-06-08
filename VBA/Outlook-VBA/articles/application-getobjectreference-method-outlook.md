---
title: Application.GetObjectReference Method (Outlook)
keywords: vbaol11.chm734
f1_keywords:
- vbaol11.chm734
ms.prod: outlook
api_name:
- Outlook.Application.GetObjectReference
ms.assetid: 426ade68-155b-9076-b3f8-4108f44688b0
ms.date: 06/08/2017
---


# Application.GetObjectReference Method (Outlook)

Creates a strong or weak object reference for a specified Outlook object.


## Syntax

 _expression_ . **GetObjectReference**( **_Item_** , **_ReferenceType_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Item_|Required| **Object**|The object from which to obtain a strong or weak object reference.|
| _ReferenceType_|Required| **[OlReferenceType](olreferencetype-enumeration-outlook.md)**|The type of object reference.|

### Return Value

An  **Object** that represents a strong or weak object reference for the specified object.


## Remarks

This method returns a weak or strong object reference for the object specified in  _Item_.


 **Note**  Outlook can fail to close successfully if an add-in retains strong object references. Always dereference a strong object reference once it is no longer needed by the add-in.


## See also


#### Concepts


[Application Object](application-object-outlook.md)

