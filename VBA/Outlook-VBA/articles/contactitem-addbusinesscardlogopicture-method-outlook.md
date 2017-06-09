---
title: ContactItem.AddBusinessCardLogoPicture Method (Outlook)
keywords: vbaol11.chm3229
f1_keywords:
- vbaol11.chm3229
ms.prod: outlook
api_name:
- Outlook.ContactItem.AddBusinessCardLogoPicture
ms.assetid: 73e19806-6892-f378-cc38-70e9d90922d1
ms.date: 06/08/2017
---


# ContactItem.AddBusinessCardLogoPicture Method (Outlook)

Adds a logo picture to the current Electronic Business Card of the contact item.


## Syntax

 _expression_ . **AddBusinessCardLogoPicture**( **_Path_** )

 _expression_ A variable that represents a **ContactItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Path_|Required| **String**|The full path name that specifies the picture file to load.|

## Remarks

An Electronic Business Card can only have one logo picture, so any existing logo pictures will be replaced. Standard graphic formats are supported, including .BMP, .GIF, .JPG, and .PNG.


## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

