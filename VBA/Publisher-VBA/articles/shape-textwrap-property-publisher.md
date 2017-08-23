---
title: "Свойство Shape.TextWrap (издатель)"
keywords: vbapb10.chm2228352
f1_keywords: vbapb10.chm2228352
ms.prod: publisher
api_name: Publisher.Shape.TextWrap
ms.assetid: e641d9a5-5b63-06d0-a0c3-d3feb1910159
ms.date: 06/08/2017
ms.openlocfilehash: 9a978b60e0200d2ab1545e42a4247a569d894119
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapetextwrap-property-publisher"></a>Свойство Shape.TextWrap (издатель)

Возвращает объект **[WrapFormat](wrapformat-object-publisher.md)** , который представляет свойства обтекания текста фигуры или диапазона фигуры.


## <a name="syntax"></a>Синтаксис

 _выражение_. **TextWrap**

 переменная _expression_A, представляющий объект **фигуры** .


## <a name="example"></a>Пример

В следующем примере добавляет овала active публикации и указывает, что текст публикации обтекания слева и справа квадрата, circumscribes овала. Будет поля 0,1 дюйма между текст публикации и верхней, нижней, левой и правой части квадрата.


```vb
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


