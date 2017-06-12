---
title: UserAccessList.Add Method (Excel)
keywords: vbaxl10.chm726075
f1_keywords:
- vbaxl10.chm726075
ms.prod: excel
api_name:
- Excel.UserAccessList.Add
ms.assetid: dd3b3bc4-8618-b680-7409-c431a12374b0
ms.date: 06/08/2017
---


# UserAccessList.Add Method (Excel)

Adds a user access list.


## Syntax

 _expression_ . **Add**( **_Name_** , **_AllowEdit_** )

 _expression_ A variable that represents an **UserAccessList** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the user access list.|
| _AllowEdit_|Required| **Boolean**| **True** allows users on the access list to edit the editable ranges on a protected worksheet.|

### Return Value

A  **[UserAccess](useraccess-object-excel.md)** object that represents the new user access list.


## See also


#### Concepts


[UserAccessList Object](useraccesslist-object-excel.md)

