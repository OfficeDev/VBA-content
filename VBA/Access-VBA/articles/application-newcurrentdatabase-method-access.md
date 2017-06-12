---
title: Application.NewCurrentDatabase Method (Access)
keywords: vbaac10.chm12612
f1_keywords:
- vbaac10.chm12612
ms.prod: access
api_name:
- Access.Application.NewCurrentDatabase
ms.assetid: 6934a77e-5fa0-7e43-e159-2ffc2a944dca
ms.date: 06/08/2017
---


# Application.NewCurrentDatabase Method (Access)

Creates a new Microsoft Access database.


## Syntax

 _expression_. **NewCurrentDatabase**( ** _filepath_**, ** _FileFormat_**, ** _Template_**, ** _SiteAddress_**, ** _ListID_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _filepath_|Required|**String**|A string expression that is the name of a new database file, including the path name and the file name extension. If your network supports it, you can also specify a network path in the following form: \\Server\Share\Folder\Filename|
| _FileFormat_|Optional|**AcNewDatabaseFormat**|A  **[AcNewDatabaseFormat](acnewdatabaseformat-enumeration-access.md)** constant that specifes the file format to use for the newly created database.|
| _Template_|Optional|**Variant**|The name of the template to be used for the new database.|
| _SiteAddress_|Optional|**String**|Uniform Resource Locator (URL) of the Windows SharePoint Services 3.0 site to link to.|
| _ListID_|Optional|**String**|Globally Unique Identifer (GUID) or the name of the Windows SharePoint Services 3.0 list to link to.|

## Remarks

You can use this method to create a new database from another application that is controlling Microsoft Access through Automation, formerly called OLE Automation. For example, you can use the  **NewCurrentDatabase** method from Microsoft Excel to create a new database in the Microsoft Access window.

The  **NewCurrentDatabase** method enables you to create a new Microsoft Access database from another application through Automation. Once you have created an instance of Microsoft Access from another application, you must also create a new database. This database opens in the Microsoft Access window.


## See also


#### Concepts


[Application Object](application-object-access.md)

