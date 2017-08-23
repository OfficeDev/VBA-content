---
title: "Свойство FillFormat.GradientStyle (издатель)"
keywords: vbapb10.chm2359556
f1_keywords: vbapb10.chm2359556
ms.prod: publisher
api_name: Publisher.FillFormat.GradientStyle
ms.assetid: 38a38de1-4ed3-7919-421f-474b0b5d7b2f
ms.date: 06/08/2017
ms.openlocfilehash: 350d46bff3af3e3edd58465dc1f9a551fd00f9b0
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fillformatgradientstyle-property-publisher"></a>Свойство FillFormat.GradientStyle (издатель)

Возвращает константу **MsoGradientStyle** , указывающее, стиль градиента для указанного заполнения. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **GradientStyle**

 переменная _expression_A, представляет собой объект- **FillFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoGradientStyle


## <a name="remarks"></a>Заметки

Используйте метод [OneColorGradient](fillformat-onecolorgradient-method-publisher.md), [PresetGradient](fillformat-presetgradient-method-publisher.md)или **[TwoColorGradient](fillformat-twocolorgradient-method-publisher.md)** Установка стиль градиента для заполнения.

Попытка получить значение этого свойства для заполнения, не имеющим градиент приводит к ошибке. Свойство **[типа](fillformat-type-property-publisher.md)** определить наличие градиентной заливки.

Значение свойства **GradientStyle** может иметь одно из ** [MsoGradientStyle](http://msdn.microsoft.com/library/1f0e723f-293c-3646-fd77-da2c8842c71f%28Office.15%29.aspx)** объявленные константы в библиотеке типов, Microsoft Office.


## <a name="example"></a>Пример

В этом примере добавляет прямоугольник active публикации и задает его стиль градиентной заливки в соответствии с именем прямоуг1 фигуры. Для обеспечения работы примера прямоуг1 должна иметь градиентной заливки.


```vb
Dim intStyle As Integer 
 
With ActiveDocument.Pages(1).Shapes 
 ' Store gradient style of rect1. 
 intStyle = .Item("rect1").Fill.GradientStyle 
 ' Add new rectangle. 
 With .AddShape(Type:=msoShapeRectangle, _ 
 Left:=0, Top:=0, Width:=40, Height:=80).Fill 
 ' Set color and gradient of new rectangle. 
 .ForeColor.RGB = RGB(128, 0, 0) 
 .OneColorGradient Style:=intStyle, _ 
 Variant:=1, Degree:=1 
 End With 
End With 

```


