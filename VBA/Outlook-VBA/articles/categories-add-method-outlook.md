---
title: Categories.Add Method (Outlook)
keywords: vbaol11.chm2437
f1_keywords:
- vbaol11.chm2437
ms.prod: outlook
api_name:
- Outlook.Categories.Add
ms.assetid: f776c2a2-1b32-f4eb-de5e-6e245a60cac2
ms.date: 06/08/2017
---


# Categories.Add Method (Outlook)

Creates a new  **[Category](category-object-outlook.md)** object and appends it to the **[Categories](categories-object-outlook.md)** collection.


## Syntax

 _expression_ . **Add**( **_Name_** , **_Color_** , **_ShortcutKey_** )

 _expression_ A variable that represents a **Categories** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the new category.|
| _Color_|Optional| **[OlCategoryColor](olcategorycolor-enumeration-outlook.md)**|The color for the new category. If no value is specified, the new category is set to the first color (as specified in the order of the  **OlCategoryColor** enumeration) that is the least used, That is, if there are unused colors, the new category is set to the first unused color in the **OlCategoryColor** enumeration. If all colors in the **OlCategoryColor** enumeration have been used, then the new category is set to the first color that is used least in the **OlCategoryColor** enumeration.|
| _ShortcutKey_|Optional| **[OlCategoryShortcutKey](olcategoryshortcutkey-enumeration-outlook.md)**|The shortcut key for the new category. If no value is specified, the default value is  **OlCategoryShortcutKeyNone** .|

### Return Value

A  **Category** object that represents the new category.


## See also


#### Concepts


[Categories Object](categories-object-outlook.md)

