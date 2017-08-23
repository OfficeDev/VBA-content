---
title: "Свойство Adjustments.Item (издатель)"
keywords: vbapb10.chm2424832
f1_keywords: vbapb10.chm2424832
ms.prod: publisher
api_name: Publisher.Adjustments.Item
ms.assetid: 9adba87a-d09d-b024-f889-4dcdab961561
ms.date: 06/08/2017
ms.openlocfilehash: 83f1febb0a511d174892d3b6b7fe600cb7120864
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="adjustmentsitem-property-publisher"></a>Свойство Adjustments.Item (издатель)

Возвращает или задает **Variant** , указывающее, корректировки значение, указанное **в качестве аргумента** . Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Элемент** ( **_Индекс_**)

 переменная _expression_A, представляющий объект **корректировки** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Индекс|Обязательное свойство.| **Integer**|Номер индекса корректировки.|

## <a name="remarks"></a>Заметки

Автофигуры, соединители и объекты WordArt может иметь до восьми корректировки.

Для линейной корректировки значение корректировка 0.0 обычно соответствует левого или верхнего края фигуры, а значение 1.0 обычно соответствует правому или нижнему краю фигуры. Тем не менее корректировки можно передать за границы фигуры для некоторых фигур. Для Радиальное корректировки значение корректировка 1.0 соответствует ширину фигуры. Угловые дополнительной настройки необходимо указать значение корректировки в градусов.

Свойство **Item** применяется только к фигуры, которые имеют корректировки.


## <a name="example"></a>Пример

В этом примере добавляется два пересечение active публикации и затем задает значение для корректировки один (только один для этого типа автофигуры) на каждом нескольких.


```vb
With ActiveDocument.Pages(1).Shapes 
 .AddShape(Type:=msoShapeCross, Left:=10, Top:=10, Width:=100, _ 
 Height:=100).Adjustments.Item(1) = 0.4 
 .AddShape(Type:=msoShapeCross, Left:=150, Top:=10, Width:=100, _ 
 Height:=100).Adjustments.Item(1) = 0.2 
End With
```

В этом примере имеет тем же, что и в предыдущем примере, даже если он не использует явно свойство **Item** .




```vb
With ActiveDocument.Pages(1).Shapes 
 .AddShape(Type:=msoShapeCross, Left:=10, Top:=10, Width:=100, _ 
 Height:=100).Adjustments(1) = 0.4 
 .AddShape(Type:=msoShapeCross, Left:=150, Top:=10, Width:=100, _ 
 Height:=100).Adjustments(1) = 0.2 
End With
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект корректировки](adjustments-object-publisher.md)

