---
title: "Свойство Shape.HasTable (издатель)"
keywords: vbapb10.chm2228321
f1_keywords: vbapb10.chm2228321
ms.prod: publisher
api_name: Publisher.Shape.HasTable
ms.assetid: 6f544d9c-00a4-3047-fbfb-6f1835bbe2c6
ms.date: 06/08/2017
ms.openlocfilehash: 2165049e0ca3152ee09fb471c900c31fc61280bf
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapehastable-property-publisher"></a>Свойство Shape.HasTable (издатель)

Возвращает **msoTrue** , если фигуры представляет объект **TableFrame** или **msoFalse** , если фигуры представляет любой другой тип объекта. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **HasTable**

 переменная _expression_A, представляющий объект **фигуры** .


## <a name="remarks"></a>Заметки

Значение свойства **HasTable** может иметь одно из **MsoTriState** константы объявляются в библиотеке типов Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**| Фигуры в диапазоне представляет объект **TableFrame** .|
| **msoTriStateMixed**|Указывает сочетание **msoTrue** и **msoFalse** для диапазона указанной фигуры.|
| **msoTrue**|Фигуры в диапазоне представляют объект **TableFrame** .|

## <a name="example"></a>Пример

В этом примере проверяется выбранной фигуре ли таблица. Если он установлен, код задает ширину столбцов один к одному дюйма (72 точки).


```vb
Sub IsTable() 
 
 With Application.Selection.ShapeRange 
 If .HasTable = msoTrue Then 
 .Table.Columns(1).Width = 72 
 End If 
 End With 
 
End Sub
```


