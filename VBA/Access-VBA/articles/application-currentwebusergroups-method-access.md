---
title: Application.CurrentWebUserGroups Method (Access)
keywords: vbaac10.chm14600
f1_keywords:
- vbaac10.chm14600
ms.prod: access
api_name:
- Access.Application.CurrentWebUserGroups
ms.assetid: efe80f7a-b6ac-12a5-3704-6e662c87e134
ms.date: 06/08/2017
---


# Application.CurrentWebUserGroups Method (Access)

Gets the collection of Microsoft SharePoint Foundation 2010 groups of which the user is a member. 


## Syntax

 _expression_. **CurrentWebUserGroups**( ** _DisplayOption_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DisplayOption_|Required|**AcWebUserGroupsDisplay**|Specifies the type of information to return about the user's groups.|

### Return Value

Variant


## Remarks

The  **CurrentWebUserGroups** method returns **Null** if the user is not a member of any groups.


## See also


#### Concepts


[Application Object](application-object-access.md)

