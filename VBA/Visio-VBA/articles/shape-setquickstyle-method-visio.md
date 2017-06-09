---
title: Shape.SetQuickStyle Method (Visio)
ms.prod: visio
ms.assetid: aebe80cb-fae9-0be7-e903-882f6eb58b63
ms.date: 06/08/2017
---


# Shape.SetQuickStyle Method (Visio)

Sets the quick style of the specified shape.


## Syntax

 _expression_ . **SetQuickStyle**_(lineMatrix,_ _fillMatrix,_ _effectsMatrix,_ _fontMatrix,_ _lineColor,_ _fillColor,_ _shadowColor,_ _fontColor)_

 _expression_ A variable that represents a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
|||||
| _lineMatrix_|Required|[VISQUICKSTYLEMATRIXINDICES](visquickstylematrixindices-enumeration-visio.md)|Specifies the shape style index that determines the line-formatting properties (for example, dash type or weight) to retrieve from the active theme and variant.|
| _fillMatrix_|Required|VISQUICKSTYLEMATRIXINDICES|Specifies the shape style index that determines the fill-formatting properties (for example, fill type or gradient stops) to retrieve from the active theme and variant.|
| _effectsMatrix_|Required|VISQUICKSTYLEMATRIXINDICES|Specifies the shape style index that determines the font-formatting properties (for example, font style) to retrieve from the active theme and variant.|
| _fontMatrix_|Required|VISQUICKSTYLEMATRIXINDICES|Specifies the shape style index that determines the font-formatting properties (for example, font style) to retrieve from the active theme and variant.|
| _lineColor_|Required|[VISQUICKSTYLECOLORS](visquickstylecolors-enumeration-visio.md)|Specifies the color index that determines the line color to retrieve from the active theme and variant to use in line formatting.|
| _fillColor_|Required|VISQUICKSTYLECOLORS|Specifies the color index that determines the fill color to retrieve from the active theme and variant to use in fill formatting.|
| _shadowColor_|Required|VISQUICKSTYLECOLORS|Specifies the color index that determines the shadow color to retrieve from the active theme and variant to use in shadows.|
| _fontColor_|Required|VISQUICKSTYLECOLORS|Specifies the color index that determines the font color to retrieve from the active theme and variant to use in shape text.|

### Return value

 **VOID**


## Remarks

Calling the  **SetQuickStyle** method corresponds to setting a quick style in the **Shape Styles** gallery on the **Home** tab of the ribbon.


## See also


#### Concepts


[Shape Object](shape-object-visio.md)

