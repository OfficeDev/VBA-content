---
title: "Свойство ThreeDFormat.PresetThreeDFormat (издатель)"
keywords: vbapb10.chm3801352
f1_keywords: vbapb10.chm3801352
ms.prod: publisher
api_name: Publisher.ThreeDFormat.PresetThreeDFormat
ms.assetid: da0b2e3e-57e5-9c6f-6d08-3f60d38ba1f8
ms.date: 06/08/2017
ms.openlocfilehash: 866cb8dbad5c09fec2939f67840afb9cc85c8663
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="threedformatpresetthreedformat-property-publisher"></a>Свойство ThreeDFormat.PresetThreeDFormat (издатель)

Возвращает константу **MsoPresetThreeDFormat** , представляющий формат предварительно придания объема. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PresetThreeDFormat**

 переменная _expression_A, представляет собой объект- **ThreeDFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoPresetThreeDFormat


## <a name="remarks"></a>Заметки

Значение свойства **PresetThreeDFormat** может иметь одно из ** [MsoPresetThreeDFormat](http://msdn.microsoft.com/library/9d362115-1979-d079-d7e5-2e7788da614b%28Office.15%29.aspx)** объявленные константы в библиотеке типов, Microsoft Office.

Каждый формат предварительно придания объема содержит набор предварительно значений для различных свойств изменяется. Если изменяется пользовательского формата, а не предварительно формата, данное свойство возвращает **msoPresetThreeDFormatMixed**. 

Значения для этого свойства соответствуют параметрам (нумеруются слева направо, сверху вниз) отображаются при нажатии кнопки **объем** на панели инструментов **Форматирование** .

Используйте метод **[SetThreeDFormat](threedformat-setthreedformat-method-publisher.md)** для задания формата предварительно придания объема.


## <a name="example"></a>Пример

В этом примере задает формат придания объема для первой фигуры на первой странице active публикации для объемных 12 стиль Если фигура изначально имеет формат настраиваемых придания объема. В данном примере для работы указанного фигуры должен быть объемной фигуры.


```vb
Sub SetPreset3D() 
 With ActiveDocument.Pages(1).Shapes(1).ThreeD 
 If .PresetThreeDFormat = msoPresetThreeDFormatMixed Then 
 .SetThreeDFormat msoThreeD12 
 End If 
 End With 
End Sub
```


