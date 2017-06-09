---
title: VisDataRecordsetAddOptions Enumeration (Visio)
keywords: vis_sdr.chm70495
f1_keywords:
- vis_sdr.chm70495
ms.prod: visio
ms.assetid: 240726a5-48cb-3034-99cf-a42967a95daf
ms.date: 06/08/2017
---


# VisDataRecordsetAddOptions Enumeration (Visio)

Constants passed to the  **DataRecordsets.Add** method, specifying characteristics of the data recordset to be added.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visDataRecordsetNoExternalDataUI**|1|Prevents data in the new data recordset from being displayed in the  **External Data** window.|
| **visDataRecordsetNoRefreshUI**|2|Prevents the data recordset from being displayed in the  **Refresh Data** dialog box.|
| **visDataRecordsetNoAdvConfig**|4|Limits the control users have of how the data recordset is refreshed in the  **Configure Refresh** dialog box for the data recordset. In particular, users cannot change the primary key or specify when shape data should be overwritten; however, users can set the refresh interval and can change the data source.|
| **visDataRecordsetDelayQuery**|8|Adds a data recordset but does not execute the command-string query until the next time you call the  **Refresh** method.|
| **visDataRecordsetDontCopyLinks**|16|Adds a data recordset, but shape-data links are not copied to the Clipboard when shapes are copied or cut.|

