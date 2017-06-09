---
title: CodeProject Object (Access)
keywords: vbaac10.chm12741
f1_keywords:
- vbaac10.chm12741
ms.prod: access
api_name:
- Access.CodeProject
ms.assetid: 70b71f57-df23-2cf7-23f5-147053a8ec26
ms.date: 06/08/2017
---


# CodeProject Object (Access)

The  **CodeProject** object refers to the project for the code database of a Microsoft Access project (.adp) or Access database.


## Remarks

The  **CodeProject** object has several collections that contain specific[AccessObject](accessobject-object-access.md)objects within the code database. The following table lists the name of each collection defined by Access project and the types of objects it contains.



|**Collections**|**Object type**|
|:-----|:-----|
|[AllForms](allforms-object-access.md)|All forms|
|[AllReports](allreports-object-access.md)|All reports|
|[AllMacros](allmacros-object-access.md)|All macros|
|[AllModules](allmodules-object-access.md)|All modules|

 **Note**   The collections in the preceding table contain all of the respective objects in the database regardless if they are opened or closed.

For example, an  **AccessObject** object representing a form is a member of the **AllForms** collection, which is a collection of **AccessObject** objects within the current database. Within the **AllForms** collection, individual members of the collection are indexed beginning with zero. You can refer to an individual **AccessObject** object in the **AllForms** collection either by referring to the form by name, or by referring to its index within the collection. If you want to refer to a specific object in the **AllForms** collection, it's better to refer to it by name because a item's collection index may change. If the object name includes a space, the name must be surrounded by brackets ([ ]).



|**Syntax**|**Example**|
|:-----|:-----|
|**AllForms** ! _formname_|AllForms!OrderForm|
|**AllForms** ![ _form name_]|AllForms![Order Form]|
|**AllForms** (" _formname_")|AllForms("OrderForm")|
|**AllForms** ( _formname_)|AllForms(0)|

## Methods



|**Name**|
|:-----|
|[AddSharedImage](codeproject-addsharedimage-method-access.md)|
|[CloseConnection](codeproject-closeconnection-method-access.md)|
|[OpenConnection](codeproject-openconnection-method-access.md)|
|[UpdateDependencyInfo](codeproject-updatedependencyinfo-method-access.md)|

## Properties



|**Name**|
|:-----|
|[AccessConnection](codeproject-accessconnection-property-access.md)|
|[AllForms](codeproject-allforms-property-access.md)|
|[AllMacros](codeproject-allmacros-property-access.md)|
|[AllModules](codeproject-allmodules-property-access.md)|
|[AllReports](codeproject-allreports-property-access.md)|
|[Application](codeproject-application-property-access.md)|
|[BaseConnectionString](codeproject-baseconnectionstring-property-access.md)|
|[Connection](codeproject-connection-property-access.md)|
|[FileFormat](codeproject-fileformat-property-access.md)|
|[FullName](codeproject-fullname-property-access.md)|
|[ImportExportSpecifications](codeproject-importexportspecifications-property-access.md)|
|[IsConnected](codeproject-isconnected-property-access.md)|
|[IsTrusted](codeproject-istrusted-property-access.md)|
|[IsWeb](codeproject-isweb-property-access.md)|
|[Name](codeproject-name-property-access.md)|
|[Parent](codeproject-parent-property-access.md)|
|[Path](codeproject-path-property-access.md)|
|[ProjectType](codeproject-projecttype-property-access.md)|
|[Properties](codeproject-properties-property-access.md)|
|[RemovePersonalInformation](codeproject-removepersonalinformation-property-access.md)|
|[Resources](codeproject-resources-property-access.md)|
|[WebSite](codeproject-website-property-access.md)|
|[IsSQLBackend](codeproject-issqlbackend-property-access.md)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
