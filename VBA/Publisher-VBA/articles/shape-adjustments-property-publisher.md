---
title: "Свойство Shape.Adjustments (издатель)"
keywords: vbapb10.chm2228273
f1_keywords: vbapb10.chm2228273
ms.prod: publisher
api_name: Publisher.Shape.Adjustments
ms.assetid: 14794cba-c671-51e3-0aac-52e885a4ba7f
ms.date: 06/08/2017
ms.openlocfilehash: b0d4f9056aaa617214961d61e70f1ec9f95fcf42
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapeadjustments-property-publisher"></a>Свойство Shape.Adjustments (издатель)

Возвращает коллекцию **[корректировки](adjustments-object-publisher.md)** , представляющий все регулировщики формы для указанного объекта **фигуры** или **ShapeRange** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **Корректировки**

 переменная _expression_A, представляющий объект **фигуры** .


## <a name="remarks"></a>Заметки

Маркерами настройки соответствуют ползунков фигуры Microsoft Publisher.


## <a name="example"></a>Пример

В этом примере количество дополнительной настройки для заданной фигуры диапазона и присваивается переменной.


```vb
Public Sub Counter() 
 
 Dim intCount as Integer 
 
 ' A Shape must be in the active publication and selected. 
 intCount = Publisher.ActiveDocument.Selection _ 
 .ShapeRange(1).Adjustments.Count 
 
End Sub
```


