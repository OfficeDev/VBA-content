---
title: "Объект LineFormat (издатель)"
keywords: vbapb10.chm3473407
f1_keywords: vbapb10.chm3473407
ms.prod: publisher
api_name: Publisher.LineFormat
ms.assetid: 9c973f5a-b2d2-78b1-24c3-350f1ba4c2ab
ms.date: 06/08/2017
ms.openlocfilehash: 4ff85b8b759d13d575e966cf125932cc183f691c
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="lineformat-object-publisher"></a>Объект LineFormat (издатель)

Представляет строку и стрелки форматирования. Для строки объект **LineFormat** содержит сведения о форматировании для строки для фигуры с границей этот объект содержит сведения о форматировании для граница фигуры.
 


## <a name="example"></a>Пример

Свойство **[строки](shape-line-property-publisher.md)** используется для возврата объекта **LineFormat** . Следующий пример добавляет синий, пунктирной линии в активный документ. Существует короткий, узкий овал на момент начала строки и long, широкий треугольник в конечной точке.
 

 

```
Sub FormatLine() 
 With ActiveDocument.Pages(1).Shapes.AddLine(BeginX:=100, _ 
 BeginY:=100, EndX:=200, EndY:=300).Line 
 .DashStyle = msoLineDashDotDot 
 .ForeColor.RGB = RGB(50, 0, 128) 
 .BeginArrowheadLength = msoArrowheadShort 
 .BeginArrowheadStyle = msoArrowheadOval 
 .BeginArrowheadWidth = msoArrowheadNarrow 
 .EndArrowheadLength = msoArrowheadLong 
 .EndArrowheadStyle = msoArrowheadTriangle 
 .EndArrowheadWidth = msoArrowheadWide 
 End With 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[PresetGradient](lineformat-presetgradient-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](lineformat-application-property-publisher.md)|
|[Цвет фона](lineformat-backcolor-property-publisher.md)|
|[BeginArrowheadLength](lineformat-beginarrowheadlength-property-publisher.md)|
|[BeginArrowheadStyle](lineformat-beginarrowheadstyle-property-publisher.md)|
|[BeginArrowheadWidth](lineformat-beginarrowheadwidth-property-publisher.md)|
|[CapStyle](lineformat-capstyle-property-publisher.md)|
|[DashStyle](lineformat-dashstyle-property-publisher.md)|
|[EndArrowheadLength](lineformat-endarrowheadlength-property-publisher.md)|
|[EndArrowheadStyle](lineformat-endarrowheadstyle-property-publisher.md)|
|[EndArrowheadWidth](lineformat-endarrowheadwidth-property-publisher.md)|
|[Цвет текста](lineformat-forecolor-property-publisher.md)|
|[GradientAngle](lineformat-gradientangle-property-publisher.md)|
|[GradientColorType](lineformat-gradientcolortype-property-publisher.md)|
|[GradientStyle](lineformat-gradientstyle-property-publisher.md)|
|[GradientVariant](lineformat-gradientvariant-property-publisher.md)|
|[InsetPen](lineformat-insetpen-property-publisher.md)|
|[JoinStyle](lineformat-joinstyle-property-publisher.md)|
|[Родительский раздел](lineformat-parent-property-publisher.md)|
|[Шаблон](lineformat-pattern-property-publisher.md)|
|[PresetGradientType](lineformat-presetgradienttype-property-publisher.md)|
|[Стиль](lineformat-style-property-publisher.md)|
|[Прозрачность](lineformat-transparency-property-publisher.md)|
|[Type](lineformat-type-property-publisher.md)|
|[Visible](lineformat-visible-property-publisher.md)|
|[Вес](lineformat-weight-property-publisher.md)|

