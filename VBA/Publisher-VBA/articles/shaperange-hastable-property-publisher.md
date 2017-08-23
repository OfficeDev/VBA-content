---
title: "Свойство ShapeRange.HasTable (издатель)"
keywords: vbapb10.chm2293857
f1_keywords: vbapb10.chm2293857
ms.prod: publisher
api_name: Publisher.ShapeRange.HasTable
ms.assetid: 71ce4980-f5b5-c94c-c29d-32b97cf771fd
ms.date: 06/08/2017
ms.openlocfilehash: 18420db0755bb08426769587c2295a2b08989fd0
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangehastable-property-publisher"></a>Свойство ShapeRange.HasTable (издатель)

Возвращает **msoTrue** , если фигуры представляет объект **TableFrame** или **msoFalse** , если фигуры представляет любой другой тип объекта. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **HasTable**

 переменная _expression_A, представляющий объект **ShapeRange** .


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


