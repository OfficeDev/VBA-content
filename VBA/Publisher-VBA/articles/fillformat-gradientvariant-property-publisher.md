---
title: "Свойство FillFormat.GradientVariant (издатель)"
keywords: vbapb10.chm2359557
f1_keywords: vbapb10.chm2359557
ms.prod: publisher
api_name: Publisher.FillFormat.GradientVariant
ms.assetid: f57224dc-e9c6-e1aa-e950-97dfd5ed483f
ms.date: 06/08/2017
ms.openlocfilehash: 06b2b3aa326f32b31913aee7fad388b4a1d518c9
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fillformatgradientvariant-property-publisher"></a>Свойство FillFormat.GradientVariant (издатель)

Возвращает значение типа **Long** , указывающее, градиентный variant для указанного заполнения. Как правило значениями являются целые числа от 1 до 4 для большинства градиентные заливки. Если стиль градиента **msoGradientFromTitle** или **msoGradientFromCenter**, данное свойство возвращает 1 или 2. Значения для этого свойства соответствуют градиента вариантов (нумерованные слева направо и сверху вниз) на вкладке **градиента** в диалоговом окне **Заливки** . Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **GradientVariant**

 переменная _expression_A, представляет собой объект- **FillFormat** .


### <a name="return-value"></a>Возвращаемое значение

Длинный


## <a name="remarks"></a>Заметки

Используйте метод **[OneColorGradient](fillformat-onecolorgradient-method-publisher.md)**, **[PresetGradient](fillformat-presetgradient-method-publisher.md)**или **[TwoColorGradient](fillformat-twocolorgradient-method-publisher.md)** установка градиента variant для заполнения.


## <a name="example"></a>Пример

В этом примере добавляет прямоугольник active публикации и задает его тип variant градиентной заливки в соответствии с именем прямоуг1 фигуры. Для обеспечения работы примера прямоуг1 должна иметь градиентной заливки.


```vb
Dim intVariant As Integer 
 
With ActiveDocument.Pages(1).Shapes 
 ' Store gradient variant of rect1. 
 intVariant = .Item("rect1").Fill.GradientVariant 
 ' Add new rectangle. 
 With .AddShape(Type:=msoShapeRectangle, _ 
 Left:=0, Top:=0, Width:=40, Height:=80).Fill 
 ' Set color and gradient of new rectangle. 
 .ForeColor.RGB = RGB(128, 0, 0) 
 .OneColorGradient Style:=msoGradientHorizontal, _ 
 Variant:=intVariant, Degree:=1 
 End With 
End With 

```


