---
title: "Метод ShapeRange.MoveOutOfTextFlow (издатель)"
keywords: vbapb10.chm2294032
f1_keywords: vbapb10.chm2294032
ms.prod: publisher
api_name: Publisher.ShapeRange.MoveOutOfTextFlow
ms.assetid: 36d6b22d-f041-6dd8-ce2c-9514ac6af5ae
ms.date: 06/08/2017
ms.openlocfilehash: df05a4c9a7d604e2222b34e9091bcd065ae5ca9d
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangemoveoutoftextflow-method-publisher"></a>Метод ShapeRange.MoveOutOfTextFlow (издатель)

Перемещает указанный встроенный фигуры вне его содержащего диапазона текста, определенные в ** [Объект TextRange](textrange-object-publisher.md)** и делает фиксированной фигуры.


## <a name="syntax"></a>Синтаксис

 _выражение_. **MoveOutOfTextFlow**

 переменная _expression_A, представляющий объект **ShapeRange** .


### <a name="return-value"></a>Возвращаемое значение

Значение Nothing


## <a name="remarks"></a>Заметки

Если фигуры перемещаемых еще не является встроенной, возвращается ошибка автоматизации.

После вызова метода **MoveOutOfTextFlow** на встроенная фигура фигуры будет сохранять свое положение на странице, но оно больше не будет встроенного.


## <a name="example"></a>Пример

В следующем примере перемещается первую фигуру встроенные, содержащиеся в диапазоне заданный текст из текста.


```vb
Dim theShape As Shape 
 
Set theShape = ActiveDocument.Pages(2).Shapes(1) _ 
 .TextFrame.TextRange.InlineShapes(1) 
 
theShape.MoveOutOfTextFlow
```


