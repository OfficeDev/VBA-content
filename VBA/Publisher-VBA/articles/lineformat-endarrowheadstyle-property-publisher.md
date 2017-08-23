---
title: "Свойство LineFormat.EndArrowheadStyle (издатель)"
keywords: vbapb10.chm3408134
f1_keywords: vbapb10.chm3408134
ms.prod: publisher
api_name: Publisher.LineFormat.EndArrowheadStyle
ms.assetid: 991354c7-3f2c-a882-74d6-1c5cd3019494
ms.date: 06/08/2017
ms.openlocfilehash: 627a11da435d5255205cebd6f7980d737af657ff
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="lineformatendarrowheadstyle-property-publisher"></a>Свойство LineFormat.EndArrowheadStyle (издатель)

Возвращает или задает константой **MsoArrowheadStyle** , указывающее, стиль стрелки в конце указанной строке. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **EndArrowheadStyle**

 переменная _expression_A, представляющий объект **LineFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoArrowheadStyle


## <a name="remarks"></a>Заметки

Свойство **[BeginArrowheadStyle](lineformat-beginarrowheadstyle-property-publisher.md)** используется для возвращения или задания стиля стрелки в начале строки.

Значение свойства **EndArrowheadStyle** может иметь одно из ** [MsoArrowheadStyle](http://msdn.microsoft.com/library/e598631e-dad9-649b-767b-99e7e7ea83da%28Office.15%29.aspx)** объявленные константы в библиотеке типов, Microsoft Office.


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


