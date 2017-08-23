---
title: "Метод FillFormat.OneColorGradient (издатель)"
keywords: vbapb10.chm2359313
f1_keywords: vbapb10.chm2359313
ms.prod: publisher
api_name: Publisher.FillFormat.OneColorGradient
ms.assetid: e4ebf7c5-41af-8227-85de-10cc08ad9f91
ms.date: 06/08/2017
ms.openlocfilehash: 3ebc57fd40c87ad473b32d516624b8b7b6df59d5
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fillformatonecolorgradient-method-publisher"></a>Метод FillFormat.OneColorGradient (издатель)

Устанавливает указанный заливки градиентной один цвет.


## <a name="syntax"></a>Синтаксис

 _выражение_. **OneColorGradient** ( **_Стиль_**, **_Variant_**, **_степень_**)

 переменная _expression_A, представляет собой объект- **FillFormat** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Стиль|Обязательное свойство.| **MsoGradientStyle**|Стиль градиента.|
|Variant|Обязательное свойство.| **Длинный**|Градиентный variant. Может быть в диапазоне от 1 до 4, соответствующий четырех вариантов на вкладке **градиента** в диалоговом окне **Заливки** . Если стиль **msoGradientFromTitle** или **msoGradientFromCenter**, этот аргумент может быть 1 или 2.|
|Степень|Обязательное свойство.| **Один**|Определяет, насколько градиента. Может быть в диапазоне от 0,0 (темный) для версии 1.0 (недоступно).|

## <a name="remarks"></a>Заметки

Параметр Style может иметь одно из **MsoGradientStyle** константы объявляются в библиотеке типов Microsoft Office и показаны в следующей таблице.



| **msoGradientDiagonalDown**|| **msoGradientDiagonalUp**|| **msoGradientFromCenter**|| **msoGradientFromCorner**|| **msoGradientFromTitle**|| **msoGradientHorizontal**|| **msoGradientVertical**|

## <a name="example"></a>Пример

В этом примере добавляется прямоугольник с одним цвет градиентной заливки active публикацию.


```vb
With ActiveDocument.Pages(1).Shapes _ 
 .AddShape(Type:=msoShapeRectangle, _ 
 Left:=90, Top:=90, Width:=90, Height:=80).Fill 
 .ForeColor.RGB = RGB(0, 128, 128) 
 .OneColorGradient Style:=msoGradientHorizontal, _ 
 Variant:=1, Degree:=1 
End With 

```


