---
title: Document.DataRecordsetAdded Event (Visio)
keywords: vis_sdr.chm10562035
f1_keywords:
- vis_sdr.chm10562035
ms.prod: visio
api_name:
- Visio.Document.DataRecordsetAdded
ms.assetid: 3ddb399d-0b28-9ec7-4059-f8d3011a98c0
ms.date: 06/08/2017
---


# Document.DataRecordsetAdded Event (Visio)

Occurs when a  **DataRecordset** object is added to a **DataRecordsets** collection.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

Private Sub  _expression_ _**DataRecordsetAdded**( **_ByVal DataRecordset As [IVDATARECORDSET]_** )

 _expression_ An expression that returns a **Document** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DataRecordset_|Required| **[IVDATARECORDSET]**|The data recordset that was added.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


