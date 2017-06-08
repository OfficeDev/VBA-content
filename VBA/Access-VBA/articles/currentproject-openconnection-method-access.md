---
title: CurrentProject.OpenConnection Method (Access)
keywords: vbaac10.chm12715
f1_keywords:
- vbaac10.chm12715
ms.prod: access
api_name:
- Access.CurrentProject.OpenConnection
ms.assetid: 37b5d50c-ddc9-97d4-2b8f-068ba2702e6d
ms.date: 06/08/2017
---


# CurrentProject.OpenConnection Method (Access)

You can use the  **OpenConnection** method to open an ADO connection to an existing Microsoft Access project (.adp) or Access database as the current Access project or database in the Microsoft Access window.


## Syntax

 _expression_. **OpenConnection**( ** _BaseConnectionString_**, ** _UserID_**, ** _Password_** )

 _expression_ A variable that represents a **CurrentProject** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _BaseConnectionString_|Optional|**Variant**|A string expression that is the base connection string of the database.|
| _UserID_|Optional|**Variant**|A string expression that is the name of the existing Access project, including the path name and the file name extension. If your network supports it, you can also specify a network path in the following form: \\Server\Share\Folder\Filename.adp|
| _Password_|Optional|**Variant**|If you don't supply the filename extension, .adp is appended to the filename. You can use this method or the  **OpenCurrentDatabase** method to open .adp files.|

## Remarks

The  **OpenConnection** method is similar to the **Open** method of an ADO **Connection** object. This method establishes the physical connection to the data source. After this method successfully completes, the connection is live, the **Connection** and **BaseConnectionString** properties are set, and the Database window or data access page should be repopulated with data from the new connection. All parameters of this method are optional. If no base connection string is supplied, then the connection is re-established using the previous base connection string (but the user must call **CloseConnection** before calling **OpenConnection** again). In the case of an Access project, the **BaseConnectionString** property can only specify the SQL Server OLE DB Provider.


## See also


#### Concepts


[CurrentProject Object](currentproject-object-access.md)

