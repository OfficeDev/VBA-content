---
title: Shape.VisualBoundingBox Method (Visio)
ms.assetid: 6a7d4622-8ba5-c598-4aaa-c6297cb4c008
ms.date: 06/08/2017
ms.prod: visio
---


# Shape.VisualBoundingBox Method (Visio)

Returns the bounding rectangle of the given shape. Introduced in Office 2016.


## Syntax

 _expression_. **VisualBoundingBox**( _Flags_,  _Flags_,  _lpr8Left_,  _lpr8Bottom_,  _lpr8Right_,  _lpr8Top_)

 _expression_ A variable that represents a **Shape** object.


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
