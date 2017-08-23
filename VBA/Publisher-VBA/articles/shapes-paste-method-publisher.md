---
title: "Метод Shapes.Paste (издатель)"
keywords: vbapb10.chm2162724
f1_keywords: vbapb10.chm2162724
ms.prod: publisher
api_name: Publisher.Shapes.Paste
ms.assetid: 435dd253-ae35-1dcf-ae5a-d7dfd40abf33
ms.date: 06/08/2017
ms.openlocfilehash: e5704340c6165cc151dbd5ba1cb844d52f50f368
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapespaste-method-publisher"></a>Метод Shapes.Paste (издатель)

Вставка фигуры или текст в буфер обмена в указанной коллекции **[фигур](shapes-object-publisher.md)** в верхней части z порядке. Каждый вставленный объект становится членом указанной коллекции **фигур** . Если в буфере обмена диапазон текста, текст будет вставлен в только что созданный **TextFrame** фигуры. Возвращает объект **[ShapeRange](shaperange-object-publisher.md)** , представляющий вставленных объектов.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Вставить**

 переменная _expression_A, представляет собой объект- **фигур** .


### <a name="return-value"></a>Возвращаемое значение

ShapeRange


## <a name="example"></a>Пример

В этом примере копирует фигуры одно по одному в активной публикации в буфер обмена и вставляет его в страница 2.


```vb
With ActiveDocument 
 .Pages(1).Shapes(1).Copy 
 .Pages(2).Shapes.Paste 
End With
```


