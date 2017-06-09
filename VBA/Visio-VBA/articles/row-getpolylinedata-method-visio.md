---
title: Row.GetPolylineData Method (Visio)
keywords: vis_sdr.chm15816690
f1_keywords:
- vis_sdr.chm15816690
ms.prod: visio
api_name:
- Visio.Row.GetPolylineData
ms.assetid: 91b7f1b4-259d-9423-0c12-271287248a74
ms.date: 06/08/2017
---


# Row.GetPolylineData Method (Visio)

Returns the points recorded in a polyline row.


## Syntax

 _expression_ . **GetPolylineData**( **_Flags_** , **_xyArray()_** )

 _expression_ A variable that represents a **Row** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Flags_|Required| **Integer**|Flags that influence the points returned.|
| _xyArray()_|Required| **Double**|Out parameter. Returns an array of alternating  _x_ and _y_ values specifying the points recorded in the row.|

### Return Value

Nothing


## Remarks

If the row's type is not  **visTagPolylineTo** , an exception is raised.

If the  **GetPolylineData** method succeeds, _xyArray()_ returns a one-dimensional array of _n_ doubles (VT_R8) indexed from 0 to _n_ - 1. The parameter _xyArray()_ is an out parameter that is allocated by the **GetPolylineData** method, which passes ownership back to the caller. The caller should eventually perform **SafeArrayDestroy** on the returned array. (Microsoft Visual Basic and Visual Basic for Applications manage this for you.)

The  _Flags_ parameter is a bitmask that specifies options for returning points. Its value should be **visGeomWHPct** , **visGeomXYLocal** , or a combination of either of those values with **visGeomExcludeLastPoint** . If neither **visGeomWHPct** nor **visGeomXYLocal** is passed as part of the _Flags_ parameter, an error will be generated.



|**Constant **|**Value **|**Description **|
|:-----|:-----|:-----|
| **visGeomExcludeLastPoint**|&;H1 |Optional. The last point of the polyline (the X and Y cells in the row) will not be included in  _xyArray()_. |
| **visGeomWHPct**|&;H10 |The values returned in  _xyArray()_ will be percentages of width/height.|
| **visGeomXYLocal**|&;H20 |The values returned in  _xyArray()_ will be local, internal units in the drawing.|

