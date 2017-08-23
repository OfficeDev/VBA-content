---
title: "Метод Shapes.BuildFreeform (издатель)"
keywords: vbapb10.chm2162723
f1_keywords: vbapb10.chm2162723
ms.prod: publisher
api_name: Publisher.Shapes.BuildFreeform
ms.assetid: ea24a9a2-e72c-beb3-b17d-161ea41fff1d
ms.date: 06/08/2017
ms.openlocfilehash: 74b5efff7969f8fce8ca39e729a1a77d1b42ad45
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapesbuildfreeform-method-publisher"></a>Метод Shapes.BuildFreeform (издатель)

Создает объект freeform. Возвращает [FreeformBuilder](freeformbuilder-object-publisher.md)объекта, что представляет произвольный как оно при построении.


## <a name="syntax"></a>Синтаксис

 _выражение_. **BuildFreeform** ( **_EditingType_**, **_X1_** **_Y1_**)

 переменная _expression_A, представляет собой объект- **фигур** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|EditingType|Обязательное свойство.| **MsoEditingType**|Указывает тип редактирования первого узла.|
|X1|Обязательное свойство.| **Variant**|Горизонтальную позицию первый узел freeform документа относительно левого верхнего угла страницы.|
|Y1|Обязательное свойство.| **Variant**|Вертикальное положение первого узла в произвольный документа относительно левого верхнего угла страницы.|

### <a name="return-value"></a>Возвращаемое значение

FreeformBuilder


## <a name="remarks"></a>Заметки

Параметр EditingType может иметь одно из следующих **MsoEditingType** константы, описанные в библиотеке типов, Microsoft Office.



| **msoEditingAuto**| Добавляет тип узла, соответствующий в сегменты подключаемого. | | **msoEditingCorner**| Добавляет узел угла. |

## <a name="example"></a>Пример

Для X1 и аргументы Y1 числовые значения вычисляются в точках; строк может быть в любой устройств, поддерживаемых Microsoft Publisher (например, «2,5 дюйма»).



Чтобы добавить фигуру сегменты, используйте метод **[AddNodes](freeformbuilder-addnodes-method-publisher.md)** . После добавления по крайней мере один сегмент фигуру, можно использовать метод [ConvertToShape](freeformbuilder-converttoshape-method-publisher.md)для преобразования объекта **FreeformBuilder** в объект **фигуры** , имеющей геометрические описания, который был определен в объекте **FreeformBuilder** .




```vb
' Add a new freeform object. 
With ActiveDocument.Shapes _ 
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


