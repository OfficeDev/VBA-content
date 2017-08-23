---
title: "Свойство WrapFormat.DistanceBottom (издатель)"
keywords: vbapb10.chm786440
f1_keywords: vbapb10.chm786440
ms.prod: publisher
api_name: Publisher.WrapFormat.DistanceBottom
ms.assetid: b8aa651a-4fe1-8b0b-f0d1-672e3f060809
ms.date: 06/08/2017
ms.openlocfilehash: f199dd870ba1d5804ea5999ca370e02c3cb2733f
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="wrapformatdistancebottom-property-publisher"></a>Свойство WrapFormat.DistanceBottom (издатель)

Если свойство **[Type](wrapformat-type-property-publisher.md)** объекта **[WrapFormat](wrapformat-object-publisher.md)** **pbWrapTypeSquare**, возвращает или задает **Variant** , представляющий расстояние (в точках) между текст документа и нижний край указанные форму. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **DistanceBottom**

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


