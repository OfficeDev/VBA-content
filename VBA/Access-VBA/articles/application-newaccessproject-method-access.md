---
title: Application.NewAccessProject Method (Access)
keywords: vbaac10.chm12580
f1_keywords:
- vbaac10.chm12580
ms.prod: access
api_name:
- Access.Application.NewAccessProject
ms.assetid: e3b3b9ef-31f8-885c-5c92-d269b824fbdb
ms.date: 06/08/2017
---


# Application.NewAccessProject Method (Access)

You can use the  **NewAccessProject** method to create and open a new Microsoft Access project (.adp) as the current Access project in the Microsoft Access window.


## Syntax

 _expression_. **NewAccessProject**( ** _filepath_**, ** _Connect_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _filepath_|Required|**String**|the name of the new Access project, including the path name and the file name extension.|
| _Connect_|Optional|**Variant**|The connection string for the Access project. See the ADO  **ConnectionString** property for details about this string.|

### Return Value

Nothing


## Remarks

The  **NewAccessProject** method enables you to create a new Access project from within Microsoft Access or another application through Automation, formally called OLE Automation. For example, you can use the **NewAccessProject** method from Microsoft Excel to create a new Access project in the Access window. Once you have created an instance of Microsoft Access from another application, you must also create a new Access project. This Access project opens in the Microsoft Access window.

If the Access project identified by  _projname_ already exists, an error occurs.

The new Access project is opened under the Admin user account .




 **Note**   To open an Access database, use the **[NewCurrentDatabase](application-newcurrentdatabase-method-access.md)** method of the **[Application](application-object-access.md)** object.


## See also


#### Concepts


[Application Object](application-object-access.md)

