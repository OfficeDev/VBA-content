---
title: "Свойство FillFormat.PresetGradientType (издатель)"
keywords: vbapb10.chm2359559
f1_keywords: vbapb10.chm2359559
ms.prod: publisher
api_name: Publisher.FillFormat.PresetGradientType
ms.assetid: 1febce3f-9925-3fec-eb78-f5167585477e
ms.date: 06/08/2017
ms.openlocfilehash: 7ae59b9c9ee536e8734de0e42e61cbdccdb32120
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fillformatpresetgradienttype-property-publisher"></a>Свойство FillFormat.PresetGradientType (издатель)

Возвращает константу **MsoPresetGradientType** , представляющий предварительно тип градиента для указанного заполнения. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PresetGradientType**

 переменная _expression_A, представляет собой объект- **FillFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoPresetGradientType


## <a name="remarks"></a>Заметки

Значение свойства **PresetGradientType** может иметь одно из ** [MsoPresetGradientType](http://msdn.microsoft.com/library/d0ee19e7-bdd3-3102-61b4-dbb17d5c0363%28Office.15%29.aspx)** объявленные константы в библиотеке типов, Microsoft Office.

Используйте метод **[PresetGradient](fillformat-presetgradient-method-publisher.md)** для задания стиля типа градиента для заполнения.


## <a name="example"></a>Пример

В этом примере изменяется заливки для первой фигуры на первой странице активная публикация тумана предварительно градиентной заливки в том случае, если он уже установлено значение тумана предварительно заданного градиента. В этом примере предполагает наличие по крайней мере один фигуры на первой странице active публикации.


```vb
Sub SetGradient() 
 With ActiveDocument.Pages(1).Shapes(1).Fill 
 If .PresetGradientType <> msoGradientFog Then 
 .PresetGradient Style:=msoGradientHorizontal, _ 
 Variant:=1, PresetGradientType:=msoGradientFog 
 End If 
 End With 
End Sub
```


