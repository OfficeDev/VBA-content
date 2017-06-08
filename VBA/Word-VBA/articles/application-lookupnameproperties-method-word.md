---
title: Application.LookupNameProperties Method (Word)
keywords: vbawd10.chm158335279
f1_keywords:
- vbawd10.chm158335279
ms.prod: word
api_name:
- Word.Application.LookupNameProperties
ms.assetid: e876b098-29c5-6302-97bb-c5f25ca4e101
ms.date: 06/08/2017
---


# Application.LookupNameProperties Method (Word)

Looks up a name in the global address book list and displays the  **Properties** dialog box, which includes information about the specified name.


## Syntax

 _expression_ . **LookupNameProperties**( **_Name_** )

 _expression_ Required. A variable that represents an **[Application](application-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|A name in the global address book.|

## Remarks

If this method finds more than one match, it displays the  **Check Names** dialog box.


## Example

This example looks up the name Don Funk in the address book and displays the  **Properties** dialog box for Don Funk.


```vb
Application.LookupNameProperties Name:="Don Funk"
```


## See also


#### Concepts


[Application Object](application-object-word.md)

