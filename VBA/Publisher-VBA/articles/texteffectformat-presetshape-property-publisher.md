---
title: "Свойство TextEffectFormat.PresetShape (издатель)"
keywords: vbapb10.chm3735815
f1_keywords: vbapb10.chm3735815
ms.prod: publisher
api_name: Publisher.TextEffectFormat.PresetShape
ms.assetid: 4e98e606-d26b-aa81-0e19-5b8535ba6df1
ms.date: 06/08/2017
ms.openlocfilehash: 4432b71df9f3e256f5934063370a5d8d548d478e
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="texteffectformatpresetshape-property-publisher"></a>Свойство TextEffectFormat.PresetShape (издатель)

Возвращает или задает константой **MsoPresetTextEffectShape** , представляющий фигуры указанного WordArt. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PresetShape**

 переменная _expression_A, представляет собой объект- **TextEffectFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoPresetTextEffectShape


## <a name="remarks"></a>Заметки

Значение свойства **PresetShape** может иметь одно из ** [MsoPresetTextEffectShape](http://msdn.microsoft.com/library/d8b21a00-54af-b2cf-6a1d-4bbaef2aac78%28Office.15%29.aspx)** объявленные константы в библиотеке типов, Microsoft Office.


## <a name="example"></a>Пример

В этом примере задается фигуры первую фигуру на первой странице активная публикация шеврон которого центр точек вниз. В этом примере для работы первой фигуры должен быть WordArt фигуры.


```vb
Sub ChangeTextEffect() 
 With ActiveDocument.Pages(1).Shapes(1) 
 If .Type = msoTextEffect Then 
 .TextEffect.PresetShape = msoTextEffectShapeChevronDown 
 End If 
 End With 
End Sub
```


