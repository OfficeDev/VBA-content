---
title: "Метод Shape.Ungroup (издатель)"
keywords: vbapb10.chm2228265
f1_keywords: vbapb10.chm2228265
ms.prod: publisher
api_name: Publisher.Shape.Ungroup
ms.assetid: 2edd16fc-d607-856f-0524-bdef1e58a9da
ms.date: 06/08/2017
ms.openlocfilehash: 133ed981aceafeb554fc2d781108fb8bffa0876a
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapeungroup-method-publisher"></a>Метод Shape.Ungroup (издатель)

Отменяет группировку указанной группы фигур или любой группы фигур в диапазоне указанные форму. Если указанный фигуры объектов OLE и рисунков, Microsoft Publisher разбить его и преобразования его разгруппировании набор фигур. (Например, электронную таблицу Microsoft Office Excel внедренных преобразуется в линии и текстовых полей.) Возвращает разгруппировании фигур в виде одного объекта **[ShapeRange](shaperange-object-publisher.md)** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **Разгруппировать**

 переменная _expression_A, представляющий объект **фигуры** .


### <a name="return-value"></a>Возвращаемое значение

ShapeRange


## <a name="remarks"></a>Заметки

С помощью этого метода на фигуры, которая не является группу или встроенная фигура, изображение или объекта OLE приводит к ошибке. Кроме того Если рисунок — это растровое изображение, JPEG, GIF или PNG (Portable Network Graphics) файла возникает ошибка.

Так как в группы фигур рассматривается как один объект, Группировка и разгруппировка фигур изменения количество элементов в коллекции **фигур** и изменяет номера индекса элементов, следующие за затронутых элементов в коллекции. Кроме того, недавно разгруппировании фигур добавляются в коллекцию **фигур** на текущей странице (или страниц) или рабочие области. В результате они могут сместиться из одного семейства сайтов в другое.


## <a name="example"></a>Пример

В этом примере Разгруппировать сгруппированные фигуры на первой странице active публикации.


```vb
Dim shpLoop As Shape 
 
For Each shpLoop In ActiveDocument.Pages(1).Shapes 
 If shpLoop.Type = pbGroup Then shpLoop.Ungroup 
Next shpLoop 

```


