---
title: "Метод FreeformBuilder.AddNodes (издатель)"
keywords: vbapb10.chm3276816
f1_keywords: vbapb10.chm3276816
ms.prod: publisher
api_name: Publisher.FreeformBuilder.AddNodes
ms.assetid: 29906bde-e6a6-f661-0f3f-085f39653e42
ms.date: 06/08/2017
ms.openlocfilehash: c0cccca2e9f9e85db8a71af294ee7bf89503c2a9
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="freeformbuilderaddnodes-method-publisher"></a>Метод FreeformBuilder.AddNodes (издатель)

Вставляет новый сегмент в конце фигуру, который создается и добавляет узлов, которые определяют сегмента. Этот метод можно использовать столько раз, сколько требуется добавить узлы freeform созданную. После добавления узлов, используйте метод **[ConvertToShape](freeformbuilder-converttoshape-method-publisher.md)** для создания freeform определенного ранее.


## <a name="syntax"></a>Синтаксис

 _выражение_. **AddNodes** ( **_SegmentType_**, **_EditingType_**, **_X1_**, **_Y1_**, **_X2_**, **_года 2_**, **_X3_**, **_года 3_**)

 переменная _expression_A, представляет собой объект- **FreeformBuilder** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|SegmentType|Обязательное свойство.| **MsoSegmentType**|Тип сегмента будет добавлена.|
|EditingType|Обязательное свойство.| **MsoEditingType**|Указывает тип редактирования новый узел. Если SegmentType **msoSegmentLine**, EditingType должен быть **msoEditingAuto**; в противном случае возникает ошибка.|
|X1|Обязательное свойство.| **Variant**|Если EditingType новый сегмент **msoEditingAuto**, этот аргумент задает расстояние по горизонтали в левом верхнем углу страницы в конечную точку новый сегмент. Если EditingType новый узел **msoEditingCorner**, этот аргумент задает расстояние по горизонтали в левом верхнем углу страницы для первой контрольной точки для нового сегмента.|
|Y1|Обязательное свойство.| **Variant**|Если EditingType новый сегмент **msoEditingAuto**, этот аргумент задает расстояние по вертикали в левом верхнем углу страницы в конечную точку новый сегмент. Если EditingType новый узел **msoEditingCorner**, этот аргумент задает расстояние по вертикали в левом верхнем углу страницы для первой контрольной точки для нового сегмента.|
|X2|Необязательный| **Variant**|Если EditingType новый сегмент **msoEditingCorner**, этот аргумент задает расстояние по горизонтали в левом верхнем углу страницы для второй контрольной точки для нового сегмента. Если EditingType новый сегмент **msoEditingAuto**, не указать значение для этого аргумента.|
|ГОДА 2|Необязательный| **Variant**|Если EditingType новый сегмент **msoEditingCorner**, этот аргумент задает расстояние по вертикали в левом верхнем углу страницы для второй контрольной точки для нового сегмента. Если EditingType новый сегмент **msoEditingAuto**, не указать значение для этого аргумента.|
|X3|Необязательный| **Variant**|Если EditingType новый сегмент **msoEditingCorner**, этот аргумент задает расстояние по горизонтали в левом верхнем углу страницы в конечную точку новый сегмент. Если EditingType новый сегмент **msoEditingAuto**, не указать значение для этого аргумента.|
|ГОДА 3|Необязательный| **Variant**|Если EditingType новый сегмент **msoEditingAuto**, этот аргумент задает расстояние по вертикали в левом верхнем углу страницы в конечную точку новый сегмент. Если EditingType новый сегмент **msoEditingAuto**, не указать значение для этого аргумента.|

## <a name="remarks"></a>Заметки

SegmentType может иметь одно из следующих констант **MsoSegmentType** .



| **msoSegmentCurve**|| **msoSegmentLine**| EditingType может иметь одно из следующих констант **MsoEditingType** .



| **msoEditingAuto**| Добавляет тип узла, соответствующий в сегменты подключаемого. | | **msoEditingCorner**| Добавляет узел угла. | Для X1 Y1, X 2 года 2, X3 и аргументы года 3 числовые значения вычисляются в точках; строк может быть в любой устройств, поддерживаемых Microsoft Publisher (например, «2,5 дюйма»).

Чтобы добавить узлы freeform после создания iit, используйте метод **[вставки](shapenodes-insert-method-publisher.md)** **[ShapeNodes](shapenodes-object-publisher.md)** семейства сайтов.


## <a name="example"></a>Пример

В этом примере добавляется freeform с четырьмя вершинами для первой страницы в активной публикации.


```vb
' Add a new freeform object. 
With ActiveDocument.Pages(1).Shapes _ 
 .BuildFreeform(EditingType:=msoEditingCorner, _ 
 X1:=100, Y1:=100) 
 
 ' Add three more nodes and close the polygon. 
 .AddNodes SegmentType:=msoSegmentCurve, _ 
 EditingType:=msoEditingCorner, _ 
 X1:=200, Y1:=200, X2:=225, Y2:=250, X3:=250, Y3:=200 
 .AddNodes SegmentType:=msoSegmentCurve, _ 
 EditingType:=msoEditingAuto, X1:=200, Y1:=100 
 .AddNodes SegmentType:=msoSegmentLine, _ 
 EditingType:=msoEditingAuto, X1:=150, Y1:=50 
 .AddNodes SegmentType:=msoSegmentLine, _ 
 EditingType:=msoEditingAuto, X1:=100, Y1:=100 
 
 ' Convert the polygon to a Shape object. 
 .ConvertToShape 
End With 

```


