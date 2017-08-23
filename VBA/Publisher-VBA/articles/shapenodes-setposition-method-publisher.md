---
title: "Метод ShapeNodes.SetPosition (издатель)"
keywords: vbapb10.chm3473428
f1_keywords: vbapb10.chm3473428
ms.prod: publisher
api_name: Publisher.ShapeNodes.SetPosition
ms.assetid: f1a3bf8c-9778-b994-9c79-55987c6fa632
ms.date: 06/08/2017
ms.openlocfilehash: 0d9f29838e7511d93204463344e3863daf48f56c
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapenodessetposition-method-publisher"></a>Метод ShapeNodes.SetPosition (издатель)

Задает положение указанного узла. В зависимости от типа редактирования узла этот метод может повлиять на положение рядом с узлами.


## <a name="syntax"></a>Синтаксис

 _выражение_. **SetPosition** ( **_Индекса_**, **_X1_** **_Y1_**)

 переменная _expression_A, представляет собой объект- **ShapeNodes** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Индекс|Обязательное свойство.| **[INT]**|Узел является позиция которого должно быть задано. Должно быть число от 1 до количества узлов в указанном фигуры; в противном случае возникает ошибка.|
|X1|Обязательное свойство.| **Variant**|Горизонтальную позицию узла относительно левого верхнего угла страницы.|
|Y1|Обязательное свойство.| **Variant**|Вертикальное положение узел относительно левого верхнего угла страницы.|

## <a name="remarks"></a>Заметки

Для X1 и аргументы Y1 числовые значения вычисляются в точках; строк может быть в любой устройств, поддерживаемых Microsoft Publisher (например, «2,5 дюйма»).


## <a name="example"></a>Пример

В этом примере перемещает второй узел в третьей фигуры в активной публикации 200 точек вправо и 300 точек. Фигуры должен быть freeform документа.


```vb
Dim arrPoints As Variant 
Dim intX As Integer 
Dim intY As Integer 
 
With ActiveDocument.Pages(1).Shapes(3).Nodes 
 arrPoints = .Item(2).Points 
 intX = arrPoints(1, 1) 
 intY = arrPoints(1, 2) 
 .SetPosition Index:=2, X1:=intX + 200, Y1:=intY + 300 
End With 

```


