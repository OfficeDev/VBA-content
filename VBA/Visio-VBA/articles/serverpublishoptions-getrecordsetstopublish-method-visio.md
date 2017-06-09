---
title: ServerPublishOptions.GetRecordsetsToPublish Method (Visio)
keywords: vis_sdr.chm17962390
f1_keywords:
- vis_sdr.chm17962390
ms.prod: visio
api_name:
- Visio.ServerPublishOptions.GetRecordsetsToPublish
ms.assetid: d0f1981d-f0ef-12dc-a0aa-562ef38a7aec
ms.date: 06/08/2017
---


# ServerPublishOptions.GetRecordsetsToPublish Method (Visio)

Returns the identifiers (IDs) of the data recordsets that are set to be published to a server.


## Syntax

 _expression_ . **GetRecordsetsToPublish**( **_PublishDataRecordsets_** , **_DataRecordsetIDs()_** )

 _expression_ A variable that represents a **[ServerPublishOptions](serverpublishoptions-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _PublishDataRecordsets_|Required| **[VisPublishDataRecordsets](vispublishdatarecordsets-enumeration-visio.md)**|Out parameter. Returns whether all, no, or selected data recordsets are set to be published. See Remarks for possible values.|
| _DataRecordsetIDs()_|Required| **Long**|Out parameter. Returns the IDs of the data recordsets that are set to be published.|

### Return Value

 **Nothing**


## Remarks

The  _PublishDataRecordsets_ parameter can be one of the following **VisPublishDataRecordsets** constants.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visPublishDataRecordsetsAll**|0|Publish all data recordsets in the document.|
| **visPublishDataRecordsetsNone**|1|Publish none of the data recordsets in the document.|
| **visPublishDataRecordsetsSelect**|2|Publish selected data recordsets.|

