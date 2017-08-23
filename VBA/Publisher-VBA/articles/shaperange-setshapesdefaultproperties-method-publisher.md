---
title: "Метод ShapeRange.SetShapesDefaultProperties (издатель)"
keywords: vbapb10.chm2293800
f1_keywords: vbapb10.chm2293800
ms.prod: publisher
api_name: Publisher.ShapeRange.SetShapesDefaultProperties
ms.assetid: 1146cbf8-6d31-9fb8-c6a4-d54b68436cbd
ms.date: 06/08/2017
ms.openlocfilehash: 84bf0ce5f974594b265153c057dbfd3e9966e5b1
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangesetshapesdefaultproperties-method-publisher"></a>Метод ShapeRange.SetShapesDefaultProperties (издатель)

Применяет форматирование для указанного фигуры или диапазона фигуры к фигуре по умолчанию. Фигуры, созданные после использования этого метода будут иметь это форматирование, примененных к ним по умолчанию.


## <a name="syntax"></a>Синтаксис

 _выражение_. **SetShapesDefaultProperties**

 переменная _expression_A, представляющий объект **ShapeRange** .


### <a name="return-value"></a>Возвращаемое значение

Значение Nothing


## <a name="remarks"></a>Заметки

Метод **SetShapesDefaultProperties** хранит два различных наборов свойств по умолчанию, другая — для объекта **Shape** ** [Свойство AutoShapeType](shape-autoshapetype-property-publisher.md)**, а другой объект **TextFrame** . Другими словами Если этот метод вызывается для автофигуры, форматирование по умолчанию этого объекта будет применяться только к новой автофигуры и не применяется для новых текстовых полей. Если этот метод вызывается для текстового поля, форматирование по умолчанию этого объекта будет применяться только к новой текстовых полей и не будет применяться к новой автофигуры.


## <a name="example"></a>Пример

В этом примере добавляет прямоугольник active публикации, форматов заливки прямоугольника, формат прямоугольника фигуру по умолчанию и затем в документ добавляется другой прямоугольник меньшего размера. Второй прямоугольник имеет же заливки как первый из них.


```vb
With ActiveDocument.Pages(1).Shapes 
 
 With .AddShape(Type:=msoShapeRectangle, _ 
 Left:=5, Top:=5, Width:=80, Height:=60) 
 With .Fill 
 .ForeColor.RGB = RGB(0, 0, 255) 
 .BackColor.RGB = RGB(0, 204, 255) 
 .Patterned Pattern:=msoPatternHorizontalBrick 
 End With 
 .SetShapesDefaultProperties 
 End With 
 
 .AddShape Type:=msoShapeRectangle, _ 
 Left:=90, Top:=90, Width:=40, Height:=30 
 
End With 

```


