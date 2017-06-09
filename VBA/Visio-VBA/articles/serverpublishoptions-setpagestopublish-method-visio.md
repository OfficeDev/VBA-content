---
title: ServerPublishOptions.SetPagesToPublish Method (Visio)
keywords: vis_sdr.chm17962375
f1_keywords:
- vis_sdr.chm17962375
ms.prod: visio
api_name:
- Visio.ServerPublishOptions.SetPagesToPublish
ms.assetid: 9d874876-e053-d6fb-04c2-8e162a0457ec
ms.date: 06/08/2017
---


# ServerPublishOptions.SetPagesToPublish Method (Visio)

Specifies the pages to publish to a server.


## Syntax

 _expression_ . **SetPagesToPublish**( **_PublishPages_** , **_NamesArray()_** , **_ Flags_** )

 _expression_ A variable that represents a **[ServerPublishOptions](serverpublishoptions-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _PublishPages_|Required| **[VisPublishPages](vispublishpages-enumeration-visio.md)**|Indicates whether all pages or selected pages are to be published. See Remarks for possible values.|
| _NamesArray()_|Required| **String**|The names of the pages to be published, if  _PublishPages_ is **visPublishPageSelect** .|
| _Flags_|Required| **[VisLangFlags](vislangflags-enumeration-visio.md)**|Indicates whether universal or local page names are specified in  _NamesArray_. See Remarks for possible values.|

### Return Value

 **Nothing**


## Remarks

The  _PublishPages_ parameter must be one of the following **VisPublishPages** constants.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visPublishPageAll**|0|Publish all pages.|
| **visPublishPageSelect**|1|Publish selected pages.|
The  _Flags_ parameter must be one of the following **VisLangFlags** constants.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visLangLocal**|0|The page name is a local name.|
| **visLangUniversal**|1|The page name is a universal name.|

