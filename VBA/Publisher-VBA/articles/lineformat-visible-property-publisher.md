---
title: "Свойство LineFormat.Visible (издатель)"
keywords: vbapb10.chm3408146
f1_keywords: vbapb10.chm3408146
ms.prod: publisher
api_name: Publisher.LineFormat.Visible
ms.assetid: 508560d2-e143-2d0d-93e7-49141e44b521
ms.date: 06/08/2017
ms.openlocfilehash: 1158f39e3fb4d865cd1ca22a32a6ce03214ef520
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="lineformatvisible-property-publisher"></a>Свойство LineFormat.Visible (издатель)

Возвращает или задает константой **MsoTriState** , указывающее, отображается ли указанный объект или форматирование, применяемое к указанным объектом. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Visible**

 переменная _expression_A, представляет собой объект- **LineFormat** .


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


