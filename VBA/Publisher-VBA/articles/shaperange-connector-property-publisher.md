---
title: "Свойство ShapeRange.Connector (издатель)"
keywords: vbapb10.chm2293813
f1_keywords: vbapb10.chm2293813
ms.prod: publisher
api_name: Publisher.ShapeRange.Connector
ms.assetid: ce05006f-38b0-c04e-4a0f-dded72dfbc10
ms.date: 06/08/2017
ms.openlocfilehash: 82deaffa82b65fc32254a8e41640d3b4e0ebbbeb
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangeconnector-property-publisher"></a>Свойство ShapeRange.Connector (издатель)

Возвращает значение, указывающее, является ли указанный фигуры соединитель **MsoTriState**. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Соединитель**

 переменная _expression_A, представляющий объект **ShapeRange** .


## <a name="remarks"></a>Заметки

Значение свойства **соединителя** может иметь одно из **MsoTriState** константы объявляются в библиотеке типов Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**|Фигура не является соединитель.|
| **msoTriStateMixed**|Возвращает значение, указывающее, сочетание **msoTrue** и **msoFalse** для диапазона указанной фигуры.|
| **msoTrue**|Фигура — это соединитель.|

## <a name="example"></a>Пример

Этот пример удаляет все соединители по одному active публикации.


```vb
Dim i As Integer 
 
With ActiveDocument.Pages(1).Shapes 
 For i = .Count To 1 Step -1 
 With .Item(i) 
 If .Connector Then .Delete 
 End With 
 Next 
End With
```


