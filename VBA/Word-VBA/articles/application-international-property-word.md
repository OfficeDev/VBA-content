---
title: Application.International Property (Word)
keywords: vbawd10.chm158335022
f1_keywords:
- vbawd10.chm158335022
ms.prod: word
api_name:
- Word.Application.International
ms.assetid: 907c2908-01a6-2a83-9968-98c21b699f4b
ms.date: 06/08/2017
---


# Application.International Property (Word)

Returns information about the current country/region and international settings. Read-only  **Variant** .


## Syntax

 _expression_ . **International**( **_Index_** )

 _expression_ Required. A variable that represents an **[Application](application-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **WdInternationalIndex**|The current country/region and/or international setting.|

## Example

This example displays the currency format in the status bar.


```
StatusBar = "Currency Format: " _ 
 &; Application.International(wdCurrencyCode)
```


## See also


#### Concepts


[Application Object](application-object-word.md)

