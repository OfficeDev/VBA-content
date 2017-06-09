---
title: WdMailMergeState Enumeration (Word)
ms.prod: word
api_name:
- Word.WdMailMergeState
ms.assetid: b8968e19-07fc-cd5f-fbf4-3204f4946f34
ms.date: 06/08/2017
---


# WdMailMergeState Enumeration (Word)

Specifies the state of a mail merge operation.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **wdDataSource**|5|A data source with no main document.|
| **wdMainAndDataSource**|2|A main document with an attached data source.|
| **wdMainAndHeader**|3|A main document with an attached header source.|
| **wdMainAndSourceAndHeader**|4|A main document with attached data source and header source.|
| **wdMainDocumentOnly**|1|A main document with no data attached.|
| **wdNormalDocument**|0|Document is not involved in a mail merge operation.|

