---
title: OLEObject.Verb Method (Excel)
keywords: vbaxl10.chm417080
f1_keywords:
- vbaxl10.chm417080
ms.prod: excel
api_name:
- Excel.OLEObject.Verb
ms.assetid: c5714863-641c-1bfd-5688-9267494fb12d
ms.date: 06/08/2017
---


# OLEObject.Verb Method (Excel)

Sends a verb to the server of the specified OLE object.


## Syntax

 _expression_ . **Verb**( **_Verb_** )

 _expression_ A variable that represents an **OLEObject** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Verb_|Optional| **[XlOLEVerb](xloleverb-enumeration-excel.md)**|The verb that the server of the OLE object should act on. If this argument is omitted, the default verb is sent. The available verbs are determined by the object's source application. Typical verbs for an OLE object are Open and Primary (represented by the  **XlOLEVerb** constants **xlOpen** and **xlPrimary** ).|

### Return Value

Variant


## Example

This example sends the default verb to the server for OLE object one on Sheet1.


```vb
Worksheets("Sheet1").OLEObjects(1).Verb
```


## See also


#### Concepts


[OLEObject Object](oleobject-object-excel.md)

