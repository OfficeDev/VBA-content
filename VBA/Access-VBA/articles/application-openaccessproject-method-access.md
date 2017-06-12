---
title: Application.OpenAccessProject Method (Access)
keywords: vbaac10.chm12581
f1_keywords:
- vbaac10.chm12581
ms.prod: access
api_name:
- Access.Application.OpenAccessProject
ms.assetid: fdc1b231-1512-cbcd-f376-935555861b38
ms.date: 06/08/2017
---


# Application.OpenAccessProject Method (Access)

You can use the  **OpenAccessProject** method to open an existing Microsoft Access project (.adp) as the current Access project in the Microsoft Access window.


## Syntax

 _expression_. **OpenAccessProject**( ** _filepath_**, ** _Exclusive_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _filepath_|Required|**String**|The name of the existing Access project, including the path name and the file name extension.|
| _Exclusive_|Optional|**Boolean**|Specifies whether you want to open the Access project in exclusive mode. The default value is  **False**, which specifies that the Access project should be opened in shared mode.|

### Return Value

Nothing


## Remarks

The  **OpenAccessProject** method enables you to open an existing project from within Microsoft Access or another application through Automation, formally called OLE Automation. For example, you can use the **OpenAccessProject** method from Microsoft Excel to open the Northwind.adp sample database in the Microsoft Access window. Once you have created an instance of Microsoft Access from another application, you must also create a new Access project or specify a particular Access project to open. This Access project opens in the Microsoft Access window.

If you have already opened a project and wish to open another project in the Microsoft Access window, you can use the  **CloseCurrentDatabase** method to close the first Access project before opening another.


 **Note**  To open an Access database, use the  **OpenCurrentDatabase** method of the **[Application](application-object-access.md)** object.


## See also


#### Concepts


[Application Object](application-object-access.md)

