---
title: "Свойство LineFormat.BeginArrowheadStyle (издатель)"
keywords: vbapb10.chm3408130
f1_keywords: vbapb10.chm3408130
ms.prod: publisher
api_name: Publisher.LineFormat.BeginArrowheadStyle
ms.assetid: 93dcf2ed-07a3-4391-dd46-2ff9cf89ef36
ms.date: 06/08/2017
ms.openlocfilehash: 58fe02d748e7f515336a751e43938615f27d0b8c
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="lineformatbeginarrowheadstyle-property-publisher"></a>Свойство LineFormat.BeginArrowheadStyle (издатель)

Возвращает или задает константой **MsoArrowheadStyle**, указывающее, стиль стрелки в начале указанной строке. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **BeginArrowheadStyle**

 переменная _expression_A, представляет собой объект- **LineFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoArrowheadStyle


## <a name="remarks"></a>Заметки

Значение свойства **BeginArrowheadStyle** может иметь одно из ** [MsoArrowheadStyle](http://msdn.microsoft.com/library/e598631e-dad9-649b-767b-99e7e7ea83da%28Office.15%29.aspx)** объявленные константы в библиотеке типов, Microsoft Office.

Свойство **[EndArrowheadStyle](lineformat-endarrowheadstyle-property-publisher.md)** используется для возвращения или задания стиля стрелки в конце строки.


## <a name="example"></a>Пример

В этом примере добавляет строку active публикации. Существует короткий, узкий овал на начальную точку строки и long, широкий треугольник в его конечной точки.


```vb
With ActiveDocument.Pages(1).Shapes _ 
 .AddLine(BeginX:=100, BeginY:=100, _ 
 EndX:=200, EndY:=300).Line 
 .BeginArrowheadLength = msoArrowheadShort 
 .BeginArrowheadStyle = msoArrowheadOval 
 .BeginArrowheadWidth = msoArrowheadNarrow 
 .EndArrowheadLength = msoArrowheadLong 
 .EndArrowheadStyle = msoArrowheadTriangle 
 .EndArrowheadWidth = msoArrowheadWide 
End With 

```


