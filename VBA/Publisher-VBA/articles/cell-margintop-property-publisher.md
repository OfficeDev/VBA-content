---
title: "Свойство Cell.MarginTop (издатель)"
keywords: vbapb10.chm5111829
f1_keywords: vbapb10.chm5111829
ms.prod: publisher
api_name: Publisher.Cell.MarginTop
ms.assetid: f408edd3-7199-b49a-817b-7b0e8461715c
ms.date: 06/08/2017
ms.openlocfilehash: b27ef86eeb0feb8342c85a24fcbb2f8db0f2789e
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="cellmargintop-property-publisher"></a>Свойство Cell.MarginTop (издатель)

Возвращает или задает **Variant** , который представляет дискового пространства (в точках) между текстом и верхнего края ячейки, текстового фрейма или страницы. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **MarginTop**

 переменная _expression_A, представляет собой объект- **ячейки** .


## <a name="example"></a>Пример

В этом примере задается полей активной публикации для двух дюйма.


```vb
Sub SetPageMargins() 
 
 With ActiveDocument.LayoutGuides 
 .MarginTop = Application.InchesToPoints(Value:=2) 
 .MarginBottom = Application.InchesToPoints(Value:=2) 
 .MarginLeft = Application.InchesToPoints(Value:=2) 
 .MarginRight = Application.InchesToPoints(Value:=2) 
 End With 
 
End Sub
```


