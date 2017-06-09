---
title: OLEFormat.DoVerb Method (Word)
keywords: vbawd10.chm154337389
f1_keywords:
- vbawd10.chm154337389
ms.prod: word
api_name:
- Word.OLEFormat.DoVerb
ms.assetid: 9ef89849-e072-24a0-3d43-fa743154b1a2
ms.date: 06/08/2017
---


# OLEFormat.DoVerb Method (Word)

Requests that an OLE object perform one of its available verbs ? the actions an OLE object takes to activate its contents.


## Syntax

 _expression_ . **DoVerb**( **_VerbIndex_** )

 _expression_ Required. A variable that represents an **[OLEFormat](oleformat-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _VerbIndex_|Optional| **Variant**|The verb that the OLE object should perform. If this argument is omitted, the default verb is sent. If the OLE object does not support the requested verb, an error will occur. Can be any  **WdOLEVerb** constant.|

## Remarks

Each OLE object supports a set of verbs that pertain to that object.


## Example

This example sends the default verb to the server for the first floating OLE object on the active document.


```vb
ActiveDocument.Shapes(1).OLEFormat.DoVerb
```


## See also


#### Concepts


[OLEFormat Object](oleformat-object-word.md)

