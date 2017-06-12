---
title: Master.VisualBoundingBox Method (Visio)
ms.assetid: 478d636f-e741-cf6b-3e16-b5faf70a9f14
ms.date: 06/08/2017
ms.prod: visio
---


# Master.VisualBoundingBox Method (Visio)

Returns the bounding rectangle of the virtual container that has all the shapes of the given master. Introduced in Office 2016.


## Syntax

 _expression_. **VisualBoundingBox**( _Flags_,  _Flags_,  _lpr8Left_,  _lpr8Bottom_,  _lpr8Right_,  _lpr8Top_)

 _expression_ A variable that represents a **Master** object.


### Parameters


|||||
|:-----|:-----|:-----|:-----|
|Name|Optional/Requires|Data Type|Description|
| _Flags_|Required|INT16|A [VisBoundingBoxArgs Enumeration (Visio)](http://msdn.microsoft.com/library/04523cbd-758f-757d-daac-30ca4676e6c2%28Office.15%29.aspx)s constant that describe the returned rectangle.|
| _lpr8Left_|Required|DOUBLE|Left position values for the virtual bounding box.|
| _lpr8Bottom_|Required|DOUBLE|Bottom position values for the virtual bounding box.|
| _lpr8Right_|Required|DOUBLE|Right position values for the virtual bounding box.|
| _lpr8Top_|Required|DOUBLE|Top position values for the virtual bounding box.|

### Return Value

 **VOID**


## See also


#### Other resources


[VisBoundingBoxArgs Enumeration (Visio)](http://msdn.microsoft.com/library/04523cbd-758f-757d-daac-30ca4676e6c2%28Office.15%29.aspx)
