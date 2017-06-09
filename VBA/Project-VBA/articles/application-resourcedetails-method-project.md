---
title: Application.ResourceDetails Method (Project)
keywords: vbapj.chm2116
f1_keywords:
- vbapj.chm2116
ms.prod: project-server
api_name:
- Project.Application.ResourceDetails
ms.assetid: 63ac7f3c-38c6-6da9-e442-373da02b63a2
ms.date: 06/08/2017
---


# Application.ResourceDetails Method (Project)

Displays the details from a MAPI-compliant address book for a resource.


## Syntax

 _expression_. **ResourceDetails**( ** _Name_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**String**|The name of a resource to locate in the address book. If the name is found, the  **Properties** dialog box for the individual is displayed. If an exact match is not found, the mail system displays the **Check Names** dialog box to allow the user to choose a valid name from the address book. If Name is omitted, the selected resource is used.|

### Return Value

 **Boolean**


## Remarks

The  **ResourceDetails** method is available only in resource views. If no e-mail profile is available, Project displays a message that explains how to create a profile.


