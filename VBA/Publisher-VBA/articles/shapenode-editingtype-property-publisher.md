---
title: "Свойство ShapeNode.EditingType (издатель)"
keywords: vbapb10.chm3539200
f1_keywords: vbapb10.chm3539200
ms.prod: publisher
api_name: Publisher.ShapeNode.EditingType
ms.assetid: f01db634-b35a-48cd-851d-418848674686
ms.date: 06/08/2017
ms.openlocfilehash: c6c1143f4261ce11cd8d300bb94af242e8f47e31
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapenodeeditingtype-property-publisher"></a>Свойство ShapeNode.EditingType (издатель)

Если указанный узел является узел, данное свойство возвращает **MsoEditingType** константу, указывающее, влияние изменений, внесенных в узел на два сегмента, подключенных к узлу. Если узел является контрольной точки для сегмент, данное свойство возвращает типа редактирования рядом с вершины. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **EditingType**

 переменная _expression_A, представляющий объект **ShapeNode** .


### <a name="return-value"></a>Возвращаемое значение

MsoEditingType


## <a name="remarks"></a>Заметки

Используйте метод **[SetEditingType](shapenodes-seteditingtype-method-publisher.md)** для задания значения этого свойства.

Значение свойства **EditingType** может иметь одно из ** [MsoEditingType](http://msdn.microsoft.com/library/5fe5c4f6-6467-c6a7-197c-ff700c384b92%28Office.15%29.aspx)** объявленные константы в библиотеке типов, Microsoft Office.


## <a name="example"></a>Пример

В этом примере изменяется все узлы углу сгладить график узлов в третьей фигуры в активной публикации. Фигуры должен быть freeform документа.


```vb
Dim intNode As Integer 
 
With ActiveDocument.Pages(1).Shapes(3).Nodes 
 For intNode = 1 to .Count 
 If .Item(intNode).EditingType = msoEditingCorner Then 
 .SetEditingType Index:=intNode, _ 
 EditingType:=msoEditingSmooth 
 End If 
 Next 
End With 

```


