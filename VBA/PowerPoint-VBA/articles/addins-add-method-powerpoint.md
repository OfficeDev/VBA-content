---
title: AddIns.Add Method (PowerPoint)
keywords: vbapp10.chm520004
f1_keywords:
- vbapp10.chm520004
ms.prod: powerpoint
api_name:
- PowerPoint.AddIns.Add
ms.assetid: e476e0dc-e82b-c460-822b-def325330514
ms.date: 06/08/2017
---


# AddIns.Add Method (PowerPoint)

Returns an  **AddIn** object that represents an add-in file added to the list of add-ins.


## Syntax

 _expression_. **Add**( **_Filename_** )

 _expression_ A variable that represents an **AddIns** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required|**String**|The full name of the file (including the path and file name extension) that contains the add-in you want to add to the list of add-ins.|

### Return Value

AddIn


## Remarks

This method doesn't load the new add-in. You must set the  **Loaded** property to load the add-in.


## See also


#### Concepts


[AddIns Object](addins-object-powerpoint.md)

