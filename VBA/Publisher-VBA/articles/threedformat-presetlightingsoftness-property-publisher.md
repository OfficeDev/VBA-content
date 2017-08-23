---
title: "Свойство ThreeDFormat.PresetLightingSoftness (издатель)"
keywords: vbapb10.chm3801350
f1_keywords: vbapb10.chm3801350
ms.prod: publisher
api_name: Publisher.ThreeDFormat.PresetLightingSoftness
ms.assetid: 8bad53c5-9d1c-367f-3f43-64691e193334
ms.date: 06/08/2017
ms.openlocfilehash: 8225dbcf423e11b16a0c4e0688bcd14250bc4a53
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="threedformatpresetlightingsoftness-property-publisher"></a>Свойство ThreeDFormat.PresetLightingSoftness (издатель)

Возвращает или задает значение константы **MsoPresetLightingSoftness** , представляющее интенсивность освещения придания объема. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PresetLightingSoftness**

 переменная _expression_A, представляет собой объект- **ThreeDFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoPresetLightingSoftness


## <a name="remarks"></a>Заметки

Значение свойства **PresetLightingSoftness** может иметь одно из ** [MsoPresetLightingSoftness](http://msdn.microsoft.com/library/da5b4779-fca6-59cd-4cfe-599c3033c5d0%28Office.15%29.aspx)** объявленные константы в библиотеке типов, Microsoft Office.


## <a name="example"></a>Пример

В этом примере задается изменяется для первой фигуры на первой странице active публикации, чтобы быть выключены яркой слева. В данном примере для работы указанного фигуры должен быть объемной фигуры.


```vb
Sub SetExtrusionLightingBrighness() 
 With ActiveDocument.Pages(1).Shapes(1).ThreeD 
 .Visible = True 
 .PresetLightingSoftness = msoLightingBright 
 .PresetLightingDirection = msoLightingLeft 
 End With 
End Sub
```


