---
title: MediaBookmarks.Add Method (PowerPoint)
keywords: vbapp10.chm730002
f1_keywords:
- vbapp10.chm730002
ms.prod: powerpoint
api_name:
- PowerPoint.MediaBookmarks.Add
ms.assetid: 2b796284-c172-9841-2af5-5f351e4acb01
ms.date: 06/08/2017
---


# MediaBookmarks.Add Method (PowerPoint)

Adds a new  **MediaBookmark** at the specified time and using the specified name.


## Syntax

 _expression_. **Add**( **_Position_**, **_Name_** )

 _expression_ A variable that represents a **MediaBookmarks** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Position_|Required|**Long**|The position of the  **MediaBookmark**.|
| _Name_|Required|**String**|The name of the  **MediaBookmark**.|

### Return Value

MediaBookmark


## Remarks

The collection is automatically re-sorted incrementally by time. This method returns an error if the bookmark already exists at that position, if the maximum number of bookmarks exceeds 512, or if the user tries to assign a name that has a length greater than 255 characters. 


## See also


#### Concepts


[MediaBookmarks Object](mediabookmarks-object-powerpoint.md)

