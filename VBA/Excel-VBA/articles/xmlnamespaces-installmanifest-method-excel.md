---
title: XmlNamespaces.InstallManifest Method (Excel)
keywords: vbaxl10.chm746078
f1_keywords:
- vbaxl10.chm746078
ms.prod: excel
api_name:
- Excel.XmlNamespaces.InstallManifest
ms.assetid: e462d627-d4d1-b3e9-4d6c-ae7ed91665ad
ms.date: 06/08/2017
---


# XmlNamespaces.InstallManifest Method (Excel)

Installs the specified XML expansion pack on the user's computer, making an XML smart document solution available to one or more users.


## Syntax

 _expression_ . **InstallManifest**( **_Path_** , **_InstallForAllUsers_** )

 _expression_ A variable that represents a **XmlNamespaces** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Path_|Required| **String**|The path and file name of the XML expansion pack.|
| _InstallForAllUsers_|Optional| **Variant**| **True** installs the XML expansion pack and makes it available to all users on a machine. **False** makes the XML expansion pack available for the current user only. Default is **False** .|

## Remarks

For security purposes, you cannot install an unsigned manifest. For more information about manifests, see the Smart Document Software Development Kit (SDK) on the Microsoft Developer Network (MSDN) Web site.


## Example

The following example installs the SimpleSample smart document solution on the user's computer and makes it available only to the current user.


 **Note**  The SimpleSample schema is included in the Smart Document Software Development Kit (SDK). For more information, see the Smart Document SDK.


```vb
Application.XMLNamespaces.InstallManifest _ 
 "http://smartdocuments/simplesample/manifest.xml"
```


## See also


#### Concepts


[XmlNamespaces Object](xmlnamespaces-object-excel.md)

