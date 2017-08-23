---
title: "Свойство FillFormat.PresetTexture (издатель)"
keywords: vbapb10.chm2359560
f1_keywords: vbapb10.chm2359560
ms.prod: publisher
api_name: Publisher.FillFormat.PresetTexture
ms.assetid: c03a9bf3-7378-e82a-9a40-650c5c96fd2a
ms.date: 06/08/2017
ms.openlocfilehash: 42978780243ffbc10174df61382418c1836a099c
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fillformatpresettexture-property-publisher"></a>Свойство FillFormat.PresetTexture (издатель)

Возвращает константу **MsoPresetTexture** , представляющий предварительно текстуры для указанного заполнения. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PresetTexture**

 переменная _expression_A, представляет собой объект- **FillFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoPresetTexture


## <a name="remarks"></a>Заметки

Значение свойства **PresetTexture** может иметь одно из ** [MsoPresetTexture](http://msdn.microsoft.com/library/fbbc897d-f5db-eb0d-20d9-f6b7e9bbcf4f%28Office.15%29.aspx)** объявленные константы в библиотеке типов, Microsoft Office.

Метод **[PresetTextured](fillformat-presettextured-method-publisher.md)** используется для указания предварительно текстуры для заполнения.


## <a name="example"></a>Пример

В этом примере добавляет прямоугольник к первой страницы в активной публикации и задает его предварительно текстуры в соответствии с, первой фигуры на странице. Для обеспечения работы примера первую фигуру должна иметь предварительно текстуры заливки.


```vb
Sub SetTexture() 
 Dim texture As MsoPresetTexture 
 With ActiveDocument.Pages(1).Shapes 
 texture = .Item(1).Fill.PresetTexture 
 With .AddShape(Type:=msoShapeRectangle, Left:=250, Top:=72, _ 
 Width:=40, Height:=80) 
 .Fill.PresetTextured PresetTexture:=texture 
 End With 
 End With 
End Sub
```


