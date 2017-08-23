---
title: "Свойство WrapFormat.DistanceLeft (издатель)"
keywords: vbapb10.chm786439
f1_keywords: vbapb10.chm786439
ms.prod: publisher
api_name: Publisher.WrapFormat.DistanceLeft
ms.assetid: 4d05ac86-f4a2-8a5e-bc7c-e303fee67e18
ms.date: 06/08/2017
ms.openlocfilehash: 5dfdabd64558f84814fa281034b1e1e382db05b5
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="wrapformatdistanceleft-property-publisher"></a>Свойство WrapFormat.DistanceLeft (издатель)

Если свойство **[Type](wrapformat-type-property-publisher.md)** объекта **[WrapFormat](wrapformat-object-publisher.md)** **pbWrapTypeSquare**, возвращает или задает **Variant** , представляющий расстояние (в точках) между текст документа и левой границей указанного фигуры. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **DistanceLeft**

 переменная _expression_A, представляет собой объект- **WrapFormat** .


## <a name="example"></a>Пример

В этом примере добавляется овала в активный документ и указывает, что текст документа обтекания слева и справа квадрата, circumscribes овала. В этом примере поля 0,1 дюйма между текст документа и верхней, нижней, левой и правой части квадрата.


```vb
Sub AddNewShape() 
 Dim shpOval As Shape 
 
 Set shpOval = ActiveDocument.Pages(1).Shapes _ 
 .AddShape(Type:=msoShapeOval, Left:=36, _ 
 Top:=36, Width:=100, Height:=35) 
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


