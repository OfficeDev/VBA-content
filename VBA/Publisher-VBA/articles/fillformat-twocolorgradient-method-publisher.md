---
title: "Метод FillFormat.TwoColorGradient (издатель)"
keywords: vbapb10.chm2359318
f1_keywords: vbapb10.chm2359318
ms.prod: publisher
api_name: Publisher.FillFormat.TwoColorGradient
ms.assetid: 7b0d1b19-a7bf-7b3d-66f4-60dfc588abfe
ms.date: 06/08/2017
ms.openlocfilehash: e34ce5501e759a9cd60971f808f86f0d6d3e97bf
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fillformattwocolorgradient-method-publisher"></a>Метод FillFormat.TwoColorGradient (издатель)

Устанавливает указанный заливки градиентом двух цветов. Цвета заливки двух задаются свойства **[ForeColor](fillformat-forecolor-property-publisher.md)** и **[BackColor](fillformat-backcolor-property-publisher.md)** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **TwoColorGradient** ( **_Стиль_**, **_вариантов_**)

 переменная _expression_A, представляет собой объект- **FillFormat** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Стиль|Обязательное свойство.| **MsoGradientStyle**|Стиль градиента.|
|Variant|Обязательное свойство.| **Длинный**|Градиентный variant. Может быть в диапазоне от 1 до 4, соответствующий четырех вариантов на вкладке **градиента** в диалоговом окне **Заливки** . Если стиль **msoGradientFromTitle** или **msoGradientFromCenter**, этот аргумент может быть 1 или 2.|

## <a name="remarks"></a>Заметки

Параметр Style может иметь одно из **MsoGradientStyle** константы объявляются в библиотеке типов Microsoft Office и показаны в следующей таблице.



| **msoGradientDiagonalDown**|| **msoGradientDiagonalUp**|| **msoGradientFromCenter**|| **msoGradientFromCorner**|| **msoGradientFromTitle**|| **msoGradientHorizontal**|| **msoGradientVertical**|

## <a name="example"></a>Пример

В этом примере добавляет прямоугольник с градиентной заливки двух цветов active публикации и задает цвет фона и переднего плана для заполнения.


```vb
With ActiveDocument.Pages(1).Shapes _ 
 .AddShape(Type:=msoShapeRectangle, _ 
 Left:=0, Top:=0, Width:=40, Height:=80).Fill 
 .ForeColor.RGB = RGB(128, 0, 0) 
 .BackColor.RGB = RGB(0, 170, 170) 
 .TwoColorGradient Style:=msoGradientHorizontal, Variant:=1 
End With 

```


