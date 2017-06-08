---
title: ServerPublishOptions.SetRecordsetsToPublish Method (Visio)
keywords: vis_sdr.chm17962385
f1_keywords:
- vis_sdr.chm17962385
ms.prod: visio
api_name:
- Visio.ServerPublishOptions.SetRecordsetsToPublish
ms.assetid: c79a8677-e4f0-9eff-9eda-72b11d0af240
ms.date: 06/08/2017
---


# ServerPublishOptions.SetRecordsetsToPublish Method (Visio)

Sets the data recordsets to be published to a server.


## Syntax

 _expression_ . **SetRecordsetsToPublish**( **_PublishDataRecordsets_** , **_DataRecordsetIDs()_** )

 _expression_ A variable that represents a **[ServerPublishOptions](serverpublishoptions-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _PublishDataRecordsets_|Required| **[VisPublishDataRecordsets](vispublishdatarecordsets-enumeration-visio.md)**|Specifies whether all, no, or selected data recordsets are to be published. See Remarks for possible values.|
| _DataRecordsetIDs()_|Required| **Long**|Specifies the identifiers of the data recordsets that are set to be published if  _PublishDataRecordsets_ is **visPublishDataRecordsetsSelect** .|

### Return Value

 **Nothing**


## Remarks

The  _PublishDataRecordsets_ parameter must be one of the following **VisPublishDataRecordsets** constants.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visPublishDataRecordsetsAll**|0|Publish all data recordsets in the document.|
| **visPublishDataRecordsetsNone**|1|Publish none of the data recordsets in the document.|
| **visPublishDataRecordsetsSelect**|2|Publish selected data recordsets.|

