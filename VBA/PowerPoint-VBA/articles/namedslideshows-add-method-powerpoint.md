---
title: NamedSlideShows.Add Method (PowerPoint)
keywords: vbapp10.chm515004
f1_keywords:
- vbapp10.chm515004
ms.prod: powerpoint
api_name:
- PowerPoint.NamedSlideShows.Add
ms.assetid: 413ea52c-95ba-8843-af72-952303328ebd
ms.date: 06/08/2017
---


# NamedSlideShows.Add Method (PowerPoint)

Creates a new named slide show and adds it to the collection of named slide shows in the specified presentation. Returns a  **[NamedSlideShow](namedslideshow-object-powerpoint.md)** object that represents the new named slide show.


## Syntax

 _expression_. **Add**( **_Name_**, **_SafeArrayOfSlideIDs_** )

 _expression_ A variable that represents a **NamedSlideShows** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**|The name of the slide show.|
| _safeArrayOfSlideIDs_|Required|**Variant**|Contains the unique slide IDs of the slides to be displayed in a slide show.|

### Return Value

NamedSlideShow


## Remarks

The name you specify when you add a named slide show is the name you use as an argument to the  **[Run](application-run-method-powerpoint.md)** method to run the named slide show.


## See also


#### Concepts


[NamedSlideShows Object](namedslideshows-object-powerpoint.md)

