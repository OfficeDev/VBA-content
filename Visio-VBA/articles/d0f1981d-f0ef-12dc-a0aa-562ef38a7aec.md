
# ServerPublishOptions.GetRecordsetsToPublish Method (Visio)

Returns the identifiers (IDs) of the data recordsets that are set to be published to a server.


## Syntax

 _expression_ . **GetRecordsetsToPublish**( **_PublishDataRecordsets_** , **_DataRecordsetIDs()_** )

 _expression_ A variable that represents a **[ServerPublishOptions](69e71212-4ca3-9fa6-6af3-8f07af540140.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _PublishDataRecordsets_|Required| **[VisPublishDataRecordsets](f0b9ad67-fabd-d0e2-74ca-61c2e6e242b9.md)**|Out parameter. Returns whether all, no, or selected data recordsets are set to be published. See Remarks for possible values.|
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
