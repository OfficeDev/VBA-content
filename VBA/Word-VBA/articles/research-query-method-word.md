---
title: Research.Query Method (Word)
keywords: vbawd10.chm201654772
f1_keywords:
- vbawd10.chm201654772
ms.prod: word
api_name:
- Word.Research.Query
ms.assetid: 416ad3f1-d2c4-4963-81c6-ba9a639c7965
ms.date: 06/08/2017
---


# Research.Query Method (Word)

Specifies a research query.


## Syntax

 _expression_ . **Query**( **_ServiceID_** , **_QueryString_** , **_QueryLanguage_** , **_UseSelection_** , **_RequeryContextXML_** , **_NewQueryContextXML_** , **_LaunchQuery_** )

 _expression_ An expression that returns a **Research** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ServiceID_|Required| **String**|Specifies a GUID that identifies the research service.|
| _QueryString_|Optional| **String**|Specifies the query string.|
| _QueryLanguage_|Optional| **[WdLanguageID](wdlanguageid-enumeration-word.md)**|Specifies the query language of the query string.|
| _UseSelection_|Optional| **Boolean**| **True** to use the current selection as the query string. This overrides the QueryString parameter if set. Default value is **False** .|
| _LaunchQuery_|Optional| **Boolean**| **True** launches the query. False displays the **Research** task pane scoped to search the specified research service.|

### Return Value

Variant


## See also


#### Concepts


[Research Object](research-object-word.md)

