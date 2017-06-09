---
title: Research.Query Method (PowerPoint)
keywords: vbapp10.chm676003
f1_keywords:
- vbapp10.chm676003
ms.prod: powerpoint
api_name:
- PowerPoint.Research.Query
ms.assetid: 21ab6e91-7719-2714-7606-883501aa94eb
ms.date: 06/08/2017
---


# Research.Query Method (PowerPoint)

Specifies a research query.


## Syntax

 _expression_. **Query**( **_ServiceID_**, **_QueryString_**, **_QueryLanguage_**, **_UseSelection_**, **_RequeryContextXML_**, **_NewQueryContextXML_**, **_LaunchQuery_** )

 _expression_ An expression that returns a **Research** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ServiceID_|Required|**String**|Specifies a GUID that identifies the research service.|
| _QueryString_|Optional|**Variant**|Specifies the query string.|
| _QueryLanguage_|Optional|**Variant**|Specifies the query language of the query string.|
| _UseSelection_|Optional|**Boolean**|**True** to use the current selection as the query string. This overrides the QueryString parameter if set. Default value is **False**.|
| _RequeryContextXML_|Optional|**Variant**|Requery context information. This allows the caller to add additional information that the service may need. This is an XML string that is placed directly under the RequeryContext element in the query.|
| _NewQueryContextXML_|Optional|**Variant**|New query context information. This allows the caller to add additional information that the service may need. This is an XML string that is placed directly under the NewQueryContext element in the query.|
| _LaunchQuery_|Optional|**Boolean**|**True** launches the query. False displays the the **Research** task pane scoped to search the specified research service.|

## See also


#### Concepts


[Research Object](research-object-powerpoint.md)

