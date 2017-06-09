---
title: DataRecordset.RefreshInterval Property (Visio)
keywords: vis_sdr.chm16460340
f1_keywords:
- vis_sdr.chm16460340
ms.prod: visio
api_name:
- Visio.DataRecordset.RefreshInterval
ms.assetid: 3d108e6e-65af-05ea-77d2-a19d96f82c1e
ms.date: 06/08/2017
---


# DataRecordset.RefreshInterval Property (Visio)

Gets or sets how often Microsoft Visio automatically refreshes the data recordset. Read/write.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **RefreshInterval**

 _expression_ An expression that returns a **DataRecordset** object.


### Return Value

Long


## Remarks

When you create a  **DataRecordset** object, its **RefreshInterval** property is set to the default, 0. This setting indicates that data is not refreshed automatically.

By setting  **RefreshInterval** to a positive **Long** value, you can specify the time in minutes between automatic refreshes. The minimum interval you can specify is one minute.

This setting corresponds to the value a user can set in the  **Configure Refresh** dialog box.


