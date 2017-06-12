---
title: ItemProperties.Add Method (Outlook)
keywords: vbaol11.chm538
f1_keywords:
- vbaol11.chm538
ms.prod: outlook
api_name:
- Outlook.ItemProperties.Add
ms.assetid: 317daeba-e34c-8458-2492-c434707fa805
ms.date: 06/08/2017
---


# ItemProperties.Add Method (Outlook)

Adds an  **ItemProperty** object to the **ItemProperties** collection.


## Syntax

 _expression_ . **Add**( **_Name_** , **_Type_** , **_AddToFolderFields_** , **_DisplayFormat_** )

 _expression_ A variable that represents an **ItemProperties** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the new item property object.|
| _Type_|Required| **[OlUserPropertyType](oluserpropertytype-enumeration-outlook.md)**|The type of the new  **ItemProperty** .|
| _AddToFolderFields_|Optional| **Variant**|Determines if the item property will be added to the folder fields.|
| _DisplayFormat_|Optional| **Variant**|Defines the format of the field as it appears in a given folder.|

## Remarks

You can create a property of a type that is defined by the  **OlUserPropertyType** enumeration, except for the following types: **olEnumeration**,  **olOutlookInternal**, and  **olSmartFrom**.


## See also


#### Concepts


[ItemProperties Object](itemproperties-object-outlook.md)

