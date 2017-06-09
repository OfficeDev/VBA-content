---
title: OLEFormat.Verb Method (Excel)
keywords: vbaxl10.chm632076
f1_keywords:
- vbaxl10.chm632076
ms.prod: excel
api_name:
- Excel.OLEFormat.Verb
ms.assetid: bf5736e8-1909-ed0a-aaab-297ccde9ffef
ms.date: 06/08/2017
---


# OLEFormat.Verb Method (Excel)

Sends a verb to the server of the specified OLE object.


## Syntax

 _expression_ . **Verb**( **_Verb_** )

 _expression_ A variable that represents an **OLEFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Verb_|Optional| **[XlOLEVerb](xloleverb-enumeration-excel.md)**|The verb that the server of the OLE object should act on. If this argument is omitted, the default verb is sent. The available verbs are determined by the object's source application. Typical verbs for an OLE object are Open and Primary (represented by the  **XlOLEVerb** constants **xlOpen** and **xlPrimary** ).|

## See also


#### Concepts


[OLEFormat Object](oleformat-object-excel.md)

