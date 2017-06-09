---
title: Application.ResourceWindowsAccount Method (Project)
keywords: vbapj.chm2394
f1_keywords:
- vbapj.chm2394
ms.prod: project-server
api_name:
- Project.Application.ResourceWindowsAccount
ms.assetid: f03e8445-10a6-d288-b6ae-9ea2eb46f532
ms.date: 06/08/2017
---


# Application.ResourceWindowsAccount Method (Project)

Sets the security identifier for Microsoft Windows authentication for the selected resource(s), based upon a Microsoft Exchange Server Address Book.


## Syntax

 _expression_. **ResourceWindowsAccount**( ** _Name_**, ** _ShowDialog_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**String**|The name of a resource for whom a security identifier is to be obtained. If an exact match is not found, the mail system will bring up the  **Check Names** dialog box for manual selection. If Name is not specified, security identifier(s) will be obtained for the selected resource(s).|
| _ShowDialog_|Optional|**Boolean**|**True** if the user is prompted to confirm adding the security identifier to the **Windows User Account** field for each resource specified with Name. The default value is **True**.|

### Return Value

 **Boolean**


## Remarks

The  **ResourceWindowsAccount** method is only available in resource views. If the optional Security Identifier field in the Address Book is blank, the **ResourceWindowsAccount** method has no effect.


