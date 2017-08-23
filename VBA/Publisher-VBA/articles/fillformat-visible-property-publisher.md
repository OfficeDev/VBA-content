---
title: "Свойство FillFormat.Visible (издатель)"
keywords: vbapb10.chm2359571
f1_keywords: vbapb10.chm2359571
ms.prod: publisher
api_name: Publisher.FillFormat.Visible
ms.assetid: 9cbb2604-6c33-de51-71f4-8c0304868cb5
ms.date: 06/08/2017
ms.openlocfilehash: c51189b31062d0d9fa61cdf857ffacaaf2cbacb3
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fillformatvisible-property-publisher"></a>Свойство FillFormat.Visible (издатель)

Возвращает или задает константой **MsoTriState** , указывающее, отображается ли указанный объект или форматирование, применяемое к указанным объектом. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Visible**

 переменная _expression_A, представляет собой объект- **FillFormat** .


## <a name="remarks"></a>Заметки

Значение свойства **Visible** может иметь одно из **MsoTriState** константы объявляются в библиотеке типов Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**|Указанный объект или форматирования не отображается.|
| **msoTriStateMixed**|Только возвращаемое значение. Диапазон указанной фигуры содержит объекты с видимым форматирования и объектов с помощью невидимой форматирования.|
| **msoTriStateToggle**| Задайте значение только. Переключает указанный объект между видимым и исчезло невидимое.|
| **msoTrue**|Указанный объект или форматирование будет отображаться.|

## <a name="example"></a>Пример

В этом примере задается горизонтального и вертикального смещения тени фигуры три на первой странице в активной публикации. 5 точек справа от фигуры и 3 точки над текстом смещения тени. Если фигура не имеет тени, этот пример добавляет в него.


```vb
With ActiveDocument.Pages(1).Shapes(3).Shadow 
 .Visible = msoTrue 
 .OffsetX = 5 
 .OffsetY = -3 
End With
```


