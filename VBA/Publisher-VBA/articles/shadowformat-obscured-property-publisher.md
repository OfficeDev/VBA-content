---
title: "Свойство ShadowFormat.Obscured (издатель)"
keywords: vbapb10.chm3670273
f1_keywords: vbapb10.chm3670273
ms.prod: publisher
api_name: Publisher.ShadowFormat.Obscured
ms.assetid: 9bc7382e-50cf-0364-6b5a-8aa46a12d8fb
ms.date: 06/08/2017
ms.openlocfilehash: 9ad81d49fdc09cab4b4bb6cfe281b9566c8bd74c
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shadowformatobscured-property-publisher"></a>Свойство ShadowFormat.Obscured (издатель)

Возвращает или задает от **MsoTriState** значение, указывающее, отображается ли теневая указанные форму заполнения и замещается фигуры. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Закрыты**

 переменная _expression_A, представляющий объект **ShadowFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoTriState


## <a name="remarks"></a>Заметки

Значение свойства **Obscured** может иметь одно из **MsoTriState** константы объявляются в библиотеке типов Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**|Теневая указанные форму не отображается заполняется в и не закрывается фигуры Если фигура имеет без заливки.|
| **msoTriStateMixed**|Возвращает значение, указывающее, сочетание **msoTrue** и **msoFalse** для диапазона указанной фигуры.|
| **msoTriStateToggle**|Задайте значение, могут переключаться между **msoTrue** и **msoFalse**.|
| **msoTrue**|Теневая указанные форму не отображается заполняется в и не закрывается фигуры Если фигура имеет без заливки.|

## <a name="example"></a>Пример

В этом примере задается горизонтального и вертикального смещения тени для трех фигуры на странице один активный публикации. 5 точек справа от фигуры и 3 точки над текстом смещения тени. Если фигура не имеет тени, этот пример добавляет в него. Тени будет заполнено и скрыт фигуры, даже если фигуры без заливки.


```vb
With ActiveDocument.Pages(1).Shapes(3).Shadow 
 .Visible = True 
 .OffsetX = 5 
 .OffsetY = -3 
 .Obscured = msoTrue 
End With
```


