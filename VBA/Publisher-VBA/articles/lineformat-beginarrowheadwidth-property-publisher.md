---
title: "Свойство LineFormat.BeginArrowheadWidth (издатель)"
keywords: vbapb10.chm3408131
f1_keywords: vbapb10.chm3408131
ms.prod: publisher
api_name: Publisher.LineFormat.BeginArrowheadWidth
ms.assetid: a752c674-1b83-b8c8-d325-b61804f5fadc
ms.date: 06/08/2017
ms.openlocfilehash: 394afde9b84375ba33f7d06eb335eb69dfee91ee
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="lineformatbeginarrowheadwidth-property-publisher"></a>Свойство LineFormat.BeginArrowheadWidth (издатель)

Возвращает или задает константой **MsoArrowheadWidth**, указывающее ширину стрелки в начале указанной строке. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **BeginArrowheadWidth**

 переменная _expression_A, представляет собой объект- **LineFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoArrowheadWidth


## <a name="remarks"></a>Заметки

Значение свойства **BeginArrowheadWidth** может иметь одно из ** [MsoArrowheadWidth](http://msdn.microsoft.com/library/7183f2e0-7431-170b-f4e7-3f8737017ed8%28Office.15%29.aspx)** объявленные константы в библиотеке типов, Microsoft Office.

Свойство **[EndArrowheadWidth](lineformat-endarrowheadwidth-property-publisher.md)** используется для возвращения или задания ширины стрелки в конце строки.


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


