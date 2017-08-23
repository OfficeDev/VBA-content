---
title: "Свойство LineFormat.EndArrowheadWidth (издатель)"
keywords: vbapb10.chm3408135
f1_keywords: vbapb10.chm3408135
ms.prod: publisher
api_name: Publisher.LineFormat.EndArrowheadWidth
ms.assetid: 20284d2d-e733-ee26-3c1c-53fd60012a75
ms.date: 06/08/2017
ms.openlocfilehash: 50683d2a2622c262bb64b601db5719c56d15cda0
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="lineformatendarrowheadwidth-property-publisher"></a>Свойство LineFormat.EndArrowheadWidth (издатель)

Возвращает или задает константой **MsoArrowheadWidth** , указывающее ширину стрелки в конце указанной строке. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **EndArrowheadWidth**

 переменная _expression_A, представляющий объект **LineFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoArrowheadWidth


## <a name="remarks"></a>Заметки

Свойство **[BeginArrowheadWidth](lineformat-beginarrowheadwidth-property-publisher.md)** используется для возвращения или задания ширины стрелки в начале строки.

Значение свойства **EndArrowheadWidth** может иметь одно из ** [MsoArrowheadWidth](http://msdn.microsoft.com/library/7183f2e0-7431-170b-f4e7-3f8737017ed8%28Office.15%29.aspx)** объявленные константы в библиотеке типов, Microsoft Office.


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


