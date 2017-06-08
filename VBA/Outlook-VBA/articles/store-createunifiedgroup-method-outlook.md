---
title: Store.CreateUnifiedGroup Method (Outlook)
keywords: vbaol11.chm3631
f1_keywords:
- vbaol11.chm3631
ms.assetid: 45f70f08-f198-22a2-79c5-26dc3247e164
ms.date: 06/08/2017
ms.prod: outlook
---


# Store.CreateUnifiedGroup Method (Outlook)

Enables a unified group to be created.


## Syntax

 _expression_ . **CreateUnifiedGroup**( _Name_,  _Name_,  _Alias_,  _Description_,  _FAutoSubscribeMembers_,  _GroupType_)

 _expression_ A variable that represents a **Store** object.


### Parameters

The  **CreateUnifiedGroup** method takes the following parameters:



| **Name**| **Data Type**| **Description**|
| **Name**|String|Name of the group.|
| **Alias**|String|Alias of the group.|
| **Description**|String|Description of the group.|
| **FAutoSubscribeMembers**|Boolean|Subscribed members of the group.|
| **GroupType**|OLUNIFIEDGROUPTYPE|Type of group: private or public.|
| **GroupSmtpAddress**|String|Smtp address for the group.|
A call to the  **CreateUnifiedGroup** method fails when: 1) the system is not online, 2) the alias already provided by the user, or 3) a server error occurs.


### Return Value

The smtp address used to create the group.


## See also


#### Concepts


[Store Object (Outlook)](store-object-outlook.md)

