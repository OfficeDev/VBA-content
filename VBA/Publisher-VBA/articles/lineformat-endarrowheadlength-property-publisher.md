---
title: "Свойство LineFormat.EndArrowheadLength (издатель)"
keywords: vbapb10.chm3408133
f1_keywords: vbapb10.chm3408133
ms.prod: publisher
api_name: Publisher.LineFormat.EndArrowheadLength
ms.assetid: 3e46e63b-54b2-edbf-0dc1-fba2c3a5d945
ms.date: 06/08/2017
ms.openlocfilehash: 59438d9d5bf0672cee3c60bd318d8e0616c61e93
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="lineformatendarrowheadlength-property-publisher"></a>Свойство LineFormat.EndArrowheadLength (издатель)

Возвращает или задает константой **MsoArrowheadLength** , указывающее длину стрелки в конце указанной строке. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **EndArrowheadLength**

 переменная _expression_A, представляющий объект **LineFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoArrowheadLength


## <a name="remarks"></a>Заметки

Свойство **[BeginArrowheadLength](lineformat-beginarrowheadlength-property-publisher.md)** используется для возвращения или задания длина стрелки в начале строки.

Значение свойства **EndArrowheadLenght** может иметь одно из ** [MsoArrowheadLength](http://msdn.microsoft.com/library/e39957f3-ffdd-17fe-dc60-1c3f8c5b14ce%28Office.15%29.aspx)** объявленные константы в библиотеке типов, Microsoft Office.


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


