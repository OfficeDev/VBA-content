---
title: "Свойство ShadowFormat.Visible (издатель)"
keywords: vbapb10.chm3670278
f1_keywords: vbapb10.chm3670278
ms.prod: publisher
api_name: Publisher.ShadowFormat.Visible
ms.assetid: aac38753-320b-7c09-548c-318c8562e393
ms.date: 06/08/2017
ms.openlocfilehash: afd6a2380f2650bbb6f3da26d3b1eb860fd2cf5c
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shadowformatvisible-property-publisher"></a>Свойство ShadowFormat.Visible (издатель)

Возвращает или задает константой **MsoTriState** , указывающее, отображается ли указанный объект или форматирование, применяемое к указанным объектом. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Visible**

 переменная _expression_A, представляет собой объект- **ShadowFormat** .


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


