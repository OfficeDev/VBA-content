---
title: "Метод FillFormat.PresetGradient (издатель)"
keywords: vbapb10.chm2359315
f1_keywords: vbapb10.chm2359315
ms.prod: publisher
api_name: Publisher.FillFormat.PresetGradient
ms.assetid: d97c4ce8-5cef-6f53-d0c8-8bcf9ab8bb80
ms.date: 06/08/2017
ms.openlocfilehash: a8dc6ceea7ad2653aeaedd5af81cbebe66fb1b24
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fillformatpresetgradient-method-publisher"></a>Метод FillFormat.PresetGradient (издатель)

Устанавливает указанный заливки предварительно заданного градиента.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PresetGradient** ( **_Стиль_**, **_Variant_**, **_PresetGradientType_**)

 переменная _expression_A, представляет собой объект- **FillFormat** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Стиль|Обязательное свойство.| **MsoGradientStyle**|Стиль градиента.|
|Variant|Обязательное свойство.| **Длинный**|Градиентный variant. Может быть в диапазоне от 1 до 4, соответствующий четырех вариантов на вкладке **градиента** в диалоговом окне **Заливки** . Если стиль **msoGradientFromTitle** или **msoGradientFromCenter**, этот аргумент может быть 1 или 2.|
|PresetGradientType|Обязательное свойство.| **MsoPresetGradientType**|Тип градиента.|

## <a name="remarks"></a>Заметки

Параметр Style может иметь одно из **MsoPresetGradientStyle** константы объявляются в библиотеке типов Microsoft Office и показаны в следующей таблице.



| **msoGradientDiagonalDown**|| **msoGradientDiagonalUp**|| **msoGradientFromCenter**|| **msoGradientFromCorner**|| **msoGradientFromTitle**|| **msoGradientHorizontal**|| **msoGradientVertical**| Параметр PresetGradientType может быть одной из констант **MsoPresetGradientType** объявлена в библиотеке типов, Microsoft Office и показаны в следующей таблице.



| **msoGradientBrass**|| **msoGradientCalmWater**|| **msoGradientChrome**|| **msoGradientChromeII**|| **msoGradientDaybreak**|| **msoGradientDesert**|| **msoGradientEarlySunset**|| **msoGradientFire**|| **msoGradientFog**|| **msoGradientGold**|| **msoGradientGoldII**|| **msoGradientHorizon**|| **msoGradientLateSunset**|| **msoGradientMahogany**|| **msoGradientMoss**|| **msoGradientNightfall**|| **msoGradientOcean**|| **msoGradientParchment**|| **msoGradientPeacock**|| **msoGradientRainbow**|| **msoGradientRainbowII**|| **msoGradientSapphire**|| **msoGradientSilver**|| **msoGradientWheat**|

## <a name="example"></a>Пример

В этом примере добавляется прямоугольник с предварительно градиентной заливки active публикацию.


```vb
ActiveDocument.Pages(1).Shapes _ 
 .AddShape(msoShapeRectangle, 90, 90, 140, 80) _ 
 .Fill.PresetGradient Style:=msoGradientHorizontal, _ 
 Variant:=1, PresetGradientType:=msoGradientBrass 

```


