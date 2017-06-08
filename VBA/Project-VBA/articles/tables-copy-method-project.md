---
title: Tables.Copy Method (Project)
keywords: vbapj.chm132701
f1_keywords:
- vbapj.chm132701
ms.prod: project-server
api_name:
- Project.Tables.Copy
ms.assetid: dfc2f25b-e60c-ef25-9e7c-2808ce76a4ba
ms.date: 06/08/2017
---


# Tables.Copy Method (Project)

Makes a copy of a group definition for the  **Tables** collection and returns a reference to the **[Table](table-object-project.md)** object.


## Syntax

 _expression_. **Copy**( ** _Source_**, ** _NewName_** )

 _expression_ A variable that represents a **Tables** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Source_|Required|**String**|The name of the table to copy.|
| _NewName_|Required|**String**|The name of the new table.|

### Return Value

 **Table**


## See also


#### Concepts


[Tables Collection Object](tables-object-project.md)
