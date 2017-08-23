---
title: "Метод Shape.SetShapesDefaultProperties (издатель)"
keywords: vbapb10.chm2228264
f1_keywords: vbapb10.chm2228264
ms.prod: publisher
api_name: Publisher.Shape.SetShapesDefaultProperties
ms.assetid: 3f7d7143-3a08-6ff4-c28e-86598212a876
ms.date: 06/08/2017
ms.openlocfilehash: 8dccad168c15ef9a7b5944c09d5969b0def944e9
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapesetshapesdefaultproperties-method-publisher"></a>Метод Shape.SetShapesDefaultProperties (издатель)

Применяет форматирование для указанного фигуры или диапазона фигуры к фигуре по умолчанию. Фигуры, созданные после использования этого метода будут иметь это форматирование, примененных к ним по умолчанию.


## <a name="syntax"></a>Синтаксис

 _выражение_. **SetShapesDefaultProperties**

 переменная _expression_A, представляющий объект **фигуры** .


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


