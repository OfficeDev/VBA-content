---
title: "Метод Shape.MoveOutOfTextFlow (издатель)"
keywords: vbapb10.chm2228357
f1_keywords: vbapb10.chm2228357
ms.prod: publisher
api_name: Publisher.Shape.MoveOutOfTextFlow
ms.assetid: 44411d6b-a627-f0c1-0576-2918f586ff0b
ms.date: 06/08/2017
ms.openlocfilehash: 6f2b2351730e644b6060aaba1367cd58d1346a30
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapemoveoutoftextflow-method-publisher"></a>Метод Shape.MoveOutOfTextFlow (издатель)

Перемещает указанный встроенный фигуры вне его содержащего диапазона текста, определенные в ** [Объект TextRange](textrange-object-publisher.md)** и делает фиксированной фигуры.


## <a name="syntax"></a>Синтаксис

 _выражение_. **MoveOutOfTextFlow**

 переменная _expression_A, представляющий объект **фигуры** .


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


