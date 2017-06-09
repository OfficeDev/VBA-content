---
title: DataRecordsets.ItemFromID Property (Visio)
keywords: vis_sdr.chm16313775
f1_keywords:
- vis_sdr.chm16313775
ms.prod: visio
api_name:
- Visio.DataRecordsets.ItemFromID
ms.assetid: 9f430e90-2c08-07a0-2c0d-c39d96405e06
ms.date: 06/08/2017
---


# DataRecordsets.ItemFromID Property (Visio)

Returns a  **DataRecordset** object from the **DataRecordsets** collection by using the unique ID of the object. Read-only.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **ItemFromID**( **_ID_** )

 _expression_ A variable that represents a **DataRecordsets** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ID_|Required| **Long**|The unique ID of the  **DataRecordset** object to retrieve.|

### Return Value

DataRecordset


## Remarks

The ID of a  **DataRecordset** object is never recycled for a particular document. You can get the ID of a **DataRecordset** object by getting the value of the **DataRecordset.ID** property.


