---
title: "Свойство Shape.Connector (издатель)"
keywords: vbapb10.chm2228277
f1_keywords: vbapb10.chm2228277
ms.prod: publisher
api_name: Publisher.Shape.Connector
ms.assetid: 6cdff1e7-59b0-9905-96f8-99b79db1acd5
ms.date: 06/08/2017
ms.openlocfilehash: b6456302f9a1fcf1c017e1bf8a7c6eab097d52ee
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapeconnector-property-publisher"></a>Свойство Shape.Connector (издатель)

Возвращает значение, указывающее, является ли указанный фигуры соединитель **MsoTriState**. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Соединитель**

 переменная _expression_A, представляющий объект **фигуры** .


## <a name="remarks"></a>Заметки

Значение свойства **соединителя** может иметь одно из **MsoTriState** константы объявляются в библиотеке типов Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**|Фигура не является соединитель.|
| **msoTriStateMixed**|Возвращает значение, указывающее, сочетание **msoTrue** и **msoFalse** для указанной фигуры.|
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


