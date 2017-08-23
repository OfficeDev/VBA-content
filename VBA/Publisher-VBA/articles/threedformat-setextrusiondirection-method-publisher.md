---
title: "Метод ThreeDFormat.SetExtrusionDirection (издатель)"
keywords: vbapb10.chm3801108
f1_keywords: vbapb10.chm3801108
ms.prod: publisher
api_name: Publisher.ThreeDFormat.SetExtrusionDirection
ms.assetid: ac01d31d-7775-8e33-3b68-6e53f952fdda
ms.date: 06/08/2017
ms.openlocfilehash: 42e7e0d793555418b4ceecd32cdf7488ae61fa39
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="threedformatsetextrusiondirection-method-publisher"></a>Метод ThreeDFormat.SetExtrusionDirection (издатель)

Направление пути очистки придания объема принимает от вытянутый фигуры.


## <a name="syntax"></a>Синтаксис

 _выражение_. **SetExtrusionDirection** ( **_PresetExtrusionDirection_**)

 переменная _expression_A, представляет собой объект- **ThreeDFormat** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|PresetExtrusionDirection|Обязательное свойство.| **MsoPresetExtrusionDirection**|Указывает направление придания объема.|

## <a name="remarks"></a>Заметки

Параметр PresetExtrusionDirection может быть одной из констант **MsoPresetExtrusionDirection** объявлена в библиотеке типов, Microsoft Office и показаны в следующей таблице.



| **msoExtrusionBottom**|| **msoExtrusionBottomLeft**|| **msoExtrusionBottomRight**|| **msoExtrusionLeft**|| **msoExtrusionNone**|| **msoExtrusionRight**|| **msoExtrusionTop**|| **msoExtrusionTopLeft**|| **msoExtrusionTopRight**| Этот метод устанавливает значение свойства [PresetExtrusionDirection](threedformat-presetextrusiondirection-property-publisher.md)направление, указанный в аргументе PresetExtrusionDirection.


## <a name="example"></a>Пример

В этом примере указывается, что изменяется для первой фигуры в активной публикации расширение к началу фигуры и что освещения для изменяется поступают из слева.


```vb
With ActiveDocument.Pages(1).Shapes(1).ThreeD 
 .Visible = True 
 .SetExtrusionDirection _ 
 PresetExtrusionDirection:=msoExtrusionTop 
 .PresetLightingDirection = msoLightingLeft 
End With 

```


