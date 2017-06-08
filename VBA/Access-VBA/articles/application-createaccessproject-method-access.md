---
title: Application.CreateAccessProject Method (Access)
keywords: vbaac10.chm12582
f1_keywords:
- vbaac10.chm12582
ms.prod: access
api_name:
- Access.Application.CreateAccessProject
ms.assetid: 66628c62-20db-e3a3-5d27-9da3846f0514
ms.date: 06/08/2017
---


# Application.CreateAccessProject Method (Access)

You can use the  **CreateAccessProject** method to create a new Microsoft Access project (.adp) on disk.


## Syntax

 _expression_. **CreateAccessProject**( ** _filepath_**, ** _Connect_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _filepath_|Required|**String**|A string expression that is the name of the new Access project, including the path name and the file name extension. If your network supports it, you can also specify a network path in the following form: \\Server\Share\Folder\Filename.adp|
| _Connect_|Optional|**Variant**|A string expression that's the valid connection string for the Access project. See the ADO  **ConnectionString** property for details about this string.|

### Return Value

Nothing


## Remarks

The  **CreateAccessProject** method enables you to create a new Access project from within Microsoft Access or another application through Automation, formally called OLE Automation. For example, you can use the **CreateAccessProject** method from Microsoft Excel to create a new Access project on disk. Once you have created an instance of Microsoft Access from another application, you must also create a new Access project.

If the Access project identified by  _projname_ already exists, an error occurs.

To create and open a new Access project as the current Access project in the Access window, use the  **[NewAccessProject](application-newaccessproject-method-access.md)** method of the **[Application](application-object-access.md)** object.

To open an existing Access project as the current Access project in the Access window, use the  **[OpenAccessProject](application-openaccessproject-method-access.md)** method of the **Application** object.


## Example

The following example creates a Microsoft Access project named "Order Entry.adp" on drive C.


```vb
Application.CreateAccessProject "C:\Order Entry.adp" 

```


## See also


#### Concepts


[Application Object](application-object-access.md)

