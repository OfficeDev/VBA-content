---
title: Research.Query Method (Excel)
keywords: vbaxl10.chm849073
f1_keywords:
- vbaxl10.chm849073
ms.prod: excel
api_name:
- Excel.Research.Query
ms.assetid: ea3b90ba-9cb4-2682-e092-6e3dd7d40aaf
ms.date: 06/08/2017
---


# Research.Query Method (Excel)

Specifies a research query.


## Syntax

 _expression_ . **Query**( **_ServiceID_** , **_QueryString_** , **_QueryLanguage_** , **_UseSelection_** , **_RequeryContextXML_** , **_NewQueryContextXML_** , **_LaunchQuery_** )

 _expression_ A variable that represents a **Research** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ServiceID_|Required| **String**|Specifies a GUID that identifies the research service.|
| _QueryString_|Optional| **Variant**|Specifies the query string.|
| _QueryLanguage_|Optional| **Variant**|Specifies the query language of the query string.|
| _UseSelection_|Optional| **Boolean**| **True** to use the current selection as the query string. This overrides the _QueryString_ parameter if set. Default value is **False** .|
| _RequeryContextXML_|Optional| **Variant**|Specifies the xml file containing the requested content.|
| _NewQueryContextXML_|Optional| **Variant**|Specifies the XML file containing the new query content.|
| _LaunchQuery_|Optional| **Boolean**| **True** launches the query. False displays the **Research** task pane scoped to search the specified research service.|

### Return Value

Variant


## See also


#### Concepts


[Research Object](research-object-excel.md)

