---
title: MasterPages.Add Method (Publisher)
keywords: vbapb10.chm589828
f1_keywords:
- vbapb10.chm589828
ms.prod: publisher
api_name:
- Publisher.MasterPages.Add
ms.assetid: af237acb-9e4c-f9d8-685c-c86d58e9ee01
ms.date: 06/08/2017
---


# MasterPages.Add Method (Publisher)

Adds a new  **Page** object to the specified **MasterPages** object and returns the new **Page** object.


## Syntax

 _expression_. **Add**( **_IsTwoPageMaster_**,  **_Abbreviation_**,  **_Description_**, )

 _expression_A variable that represents a  **MasterPages** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|IsTwoPageMaster|Optional| **Boolean**| **True** if the master page will be part of a two page spread.|
|Abbreviation|Optional| **String**|The abbreviation, or short name, for the master page. An error occurs if this is not unique.|
|Description|Optional| **String**|The description for the master page.|

### Return Value

Page


## Example

The following example adds a new master page to the active document.


```vb
ActiveDocument.MasterPages.Add _ 
 IsTwoPageMaster:=False, _ 
 Abbreviation:="X", _ 
 Description:="Master Page X" 

```


