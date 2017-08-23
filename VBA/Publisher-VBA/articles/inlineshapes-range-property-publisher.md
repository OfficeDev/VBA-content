---
title: "Свойство InlineShapes.Range (издатель)"
keywords: vbapb10.chm5767173
f1_keywords: vbapb10.chm5767173
ms.prod: publisher
api_name: Publisher.InlineShapes.Range
ms.assetid: 375843c1-5198-6981-2e7c-8abd1d0e9dff
ms.date: 06/08/2017
ms.openlocfilehash: dbea2b16cb9d2d1b04f9561853972fe4c9fe9850
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="inlineshapesrange-property-publisher"></a>Свойство InlineShapes.Range (издатель)

Возвращает коллекцию **[ShapeRange](shaperange-object-publisher.md)** , который представляет набор встроенных фигур в коллекции **InlineShapes** вызван метод которого. Это позволяет разное форматирование автономные фигур. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Диапазон** ( **_Индекс_**)

 переменная _expression_A, представляющий объект **InlineShapes** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Индекс|Необязательный| **Длинный**|Позиция индекса встроенного фигуры в коллекции **ShapeRange** .|

## <a name="example"></a>Пример

В следующем примере выполняется поиск по каждой фигуры на первой странице публикации, а также для всех встроенных фигур в рамках каждой фигуры, находит первую фигуру встроенного в пределах диапазона встроенных фигур и зеркальное отражение по вертикали.


```vb
Dim theShape As Shape 
Dim theShapes As Shapes 
 
Set theShapes = ActiveDocument.Pages(1).Shapes 
 
For Each theShape In theShapes 
 With theShape.TextFrame.TextRange 
 .InlineShapes.Range(1).Flip (msoFlipVertical) 
 End With 
Next
```


