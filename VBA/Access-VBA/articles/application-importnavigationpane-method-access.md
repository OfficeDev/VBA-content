---
title: Application.ImportNavigationPane Method (Access)
keywords: vbaac10.chm12619
f1_keywords:
- vbaac10.chm12619
ms.prod: access
api_name:
- Access.Application.ImportNavigationPane
ms.assetid: 5365ece3-e2da-031c-4e28-89115d48acf8
ms.date: 06/08/2017
---


# Application.ImportNavigationPane Method (Access)

Loads a saved Navigation Pane configuration from disk.


## Syntax

 _expression_. **ImportNavigationPane**( ** _Path_**, ** _fAppendOnly_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Path_|Required|**String**|The path and name of the XML file that contains the Navigation Pane configuration to load. |
| _fAppendOnly_|Optional|**Boolean**|Set to  **True** to append the imported categories to the existing categories. The default value is **False**.|

## See also


#### Concepts


[Application Object](application-object-access.md)

