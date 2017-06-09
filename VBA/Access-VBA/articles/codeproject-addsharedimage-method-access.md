---
title: CodeProject.AddSharedImage Method (Access)
keywords: vbaac10.chm14660
f1_keywords:
- vbaac10.chm14660
ms.prod: access
api_name:
- Access.CodeProject.AddSharedImage
ms.assetid: 7e1e0455-65e0-820e-e25c-17989a40000b
ms.date: 06/08/2017
---


# CodeProject.AddSharedImage Method (Access)

Imports the specified image into the database and adds it to the  **[SharedResources](sharedresources-object-access.md)** collection.


## Syntax

 _expression_. **AddSharedImage**( ** _SharedImageName_**, ** _FileName_** )

 _expression_ A variable that represents a **CodeProject** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SharedImageName_|Required|**String**|Specifies the string used to identify the image in the collection.|
| _FileName_|Required|**String**|Specifies the full name and path to the image file.|

## Remarks

Use the  **AddSharedImage** method when you have an image that you want to use repeatedly, such as a companny logo. The **AddSharedImage** method makes the image available in the **Insert Image** dropdown of the **Controls** group in the **Design** tab.


## See also


#### Concepts


[CodeProject Object](codeproject-object-access.md)

