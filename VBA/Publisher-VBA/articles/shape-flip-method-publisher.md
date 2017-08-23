---
title: "Метод Shape.Flip (издатель)"
keywords: vbapb10.chm2228245
f1_keywords: vbapb10.chm2228245
ms.prod: publisher
api_name: Publisher.Shape.Flip
ms.assetid: 6d0004a5-2d76-955a-64ff-140dfbc313f3
ms.date: 06/08/2017
ms.openlocfilehash: 1b6bca97aec77558f1855093c7b292ed4b8aac35
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapeflip-method-publisher"></a>Метод Shape.Flip (издатель)

Зеркальное отражение указанного фигуры вокруг оси горизонтальный или вертикальный или зеркальное отражение всех фигур в диапазоне указанного фигуры относительно горизонтальной или вертикальной оси.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Зеркальное отражение** ( **_FlipCmd_**)

 переменная _expression_A, представляющий объект **фигуры** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|FlipCmd|Обязательное свойство.| **MsoFlipCmd**| Указывает, является ли фигура зеркально по горизонтали или по вертикали.|

### <a name="return-value"></a>Возвращаемое значение

Значение Nothing


## <a name="remarks"></a>Заметки

Параметр FlipCmd может иметь одно из следующих **MsoFlipCmd** константы, описанные в библиотеке типов, Microsoft Office.



| **msoFlipHorizontal**|| **msoFlipVertical**|

## <a name="example"></a>Пример

В этом примере добавляет треугольник в первой страницы публикации active, дублирует треугольник Вертикальное зеркальное отражение повторяющихся треугольник и делает красным.


```vb
With ActiveDocument.Pages(1).Shapes _ 
 .AddShape(Type:=msoShapeRightTriangle, _ 
 Left:=10, Top:=10, Width:=50, Height:=50) _ 
 .Duplicate 
 .Fill.ForeColor.RGB = RGB(255, 0, 0) 
 .Flip msoFlipVertical 
End With 

```


