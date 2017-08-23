---
title: "Свойство LineFormat.BeginArrowheadLength (издатель)"
keywords: vbapb10.chm3408129
f1_keywords: vbapb10.chm3408129
ms.prod: publisher
api_name: Publisher.LineFormat.BeginArrowheadLength
ms.assetid: 87daaecf-3b2b-7f21-47fd-bdf192dcac60
ms.date: 06/08/2017
ms.openlocfilehash: ccfc37879e17a399849956a333ae45fa96f97715
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="lineformatbeginarrowheadlength-property-publisher"></a>Свойство LineFormat.BeginArrowheadLength (издатель)

Возвращает или задает константой **MsoArrowheadLength**, указывающее длину стрелки в начале указанной строке. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **BeginArrowheadLength**

 переменная _expression_A, представляет собой объект- **LineFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoArrowheadLength


## <a name="remarks"></a>Заметки

Значение свойства **BeginArrowheadLength** может иметь одно из ** [MsoArrowheadLength](http://msdn.microsoft.com/library/e39957f3-ffdd-17fe-dc60-1c3f8c5b14ce%28Office.15%29.aspx)** объявленные константы в библиотеке типов, Microsoft Office.

Свойство **[EndArrowheadLength](lineformat-endarrowheadlength-property-publisher.md)** используется для возвращения или задания длина стрелки в конце строки.


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


