---
title: "Объект WrapFormat (издатель)"
keywords: vbapb10.chm851967
f1_keywords: vbapb10.chm851967
ms.prod: publisher
api_name: Publisher.WrapFormat
ms.assetid: b6f80d40-2043-6944-3ed8-f26635c7fa4d
ms.date: 06/08/2017
ms.openlocfilehash: dc2bd25c97edb086c624d19150cd92069995c3fa
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="wrapformat-object-publisher"></a>Объект WrapFormat (издатель)

Представляет все свойства обтекания текста фигуры или диапазона фигуры.
 


## <a name="example"></a>Пример

Свойство **[TextWrap](shape-textwrap-property-publisher.md)** используется для возврата объекта **WrapFormat** . В следующем примере добавляет овала active публикации и указывает, что текст публикации обтекания слева и справа квадрата, circumscribes овала. Будет поля 0,1 дюйма между текст публикации и верхней, нижней, левой и правой части квадрата.
 

 

```
Sub SetTextWrapFormatProperties() 
 Dim shpOval As Shape 
 
 Set shpOval = ActiveDocument.Pages(1).Shapes.AddShape(Type:=msoShapeOval, _ 
 Left:=36, Top:=36, Width:=100, Height:=35) 
 With shpOval.TextWrap 
 .Type = pbWrapTypeSquare 
 .Side = pbWrapSideBoth 
 .DistanceAuto = msoFalse 
 .DistanceTop = InchesToPoints(0.1) 
 .DistanceBottom = InchesToPoints(0.1) 
 .DistanceLeft = InchesToPoints(0.1) 
 .DistanceRight = InchesToPoints(0.1) 
 End With 
End Sub
```


## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](wrapformat-application-property-publisher.md)|
|[DistanceAuto](wrapformat-distanceauto-property-publisher.md)|
|[DistanceBottom](wrapformat-distancebottom-property-publisher.md)|
|[DistanceLeft](wrapformat-distanceleft-property-publisher.md)|
|[DistanceRight](wrapformat-distanceright-property-publisher.md)|
|[DistanceTop](wrapformat-distancetop-property-publisher.md)|
|[Родительский раздел](wrapformat-parent-property-publisher.md)|
|[Со стороны](wrapformat-side-property-publisher.md)|
|[Type](wrapformat-type-property-publisher.md)|

