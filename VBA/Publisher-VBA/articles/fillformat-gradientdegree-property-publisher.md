---
title: "Свойство FillFormat.GradientDegree (издатель)"
keywords: vbapb10.chm2359555
f1_keywords: vbapb10.chm2359555
ms.prod: publisher
api_name: Publisher.FillFormat.GradientDegree
ms.assetid: b2eba471-5f03-4904-f876-edea4d40a908
ms.date: 06/08/2017
ms.openlocfilehash: 6127024ef9aa0cb8f83ecf320d7244d7ff930cd6
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fillformatgradientdegree-property-publisher"></a>Свойство FillFormat.GradientDegree (издатель)

Возвращает значение **одного** , указывающее, как темный или светлый цвет один градиентной заливки. Значение 0 (ноль) означает, что черный смешанный с фигуры цвет переднего плана для формирования градиента; значение 1 означает, что Технический смешанной и значений от 0 до 1 означает, что в смешанном тени более темные или более светлый цвет переднего плана. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **GradientDegree**

 переменная _expression_A, представляет собой объект- **FillFormat** .


### <a name="return-value"></a>Возвращаемое значение

Один


## <a name="remarks"></a>Заметки

Используйте метод **[OneColorGradient](fillformat-onecolorgradient-method-publisher.md)** установка градиента степени для заполнения.


## <a name="example"></a>Пример

В этом примере добавляет прямоугольник active публикации и задает степень его градиентной заливки в соответствии с именем 2 прямоугольника фигуры. Если прямоугольник 2 не имеет один цвет градиентной заливки, в этом примере приводит к ошибке.


```vb
Dim sngDegree As Single 
 
With ActiveDocument.Pages(1).Shapes 
 ' Store degree of one-color gradient. 
 sngDegree = .Item("Rectangle 2").Fill.GradientDegree 
 ' Add new rectangle. 
 With .AddShape(msoShapeRectangle, 0, 0, 40, 80).Fill 
 ' Set color and gradient for new rectangle. 
 .ForeColor.RGB = RGB(128, 0, 0) 
 .OneColorGradient Style:=msoGradientHorizontal, _ 
 Variant:=1, Degree:=sngDegree 
 End With 
End With 

```


