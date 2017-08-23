---
title: "Метод ShapeRange.Apply (издатель)"
keywords: vbapb10.chm2293776
f1_keywords: vbapb10.chm2293776
ms.prod: publisher
api_name: Publisher.ShapeRange.Apply
ms.assetid: 3531d0aa-479e-2d50-5e1e-a35f7c1e7ba6
ms.date: 06/08/2017
ms.openlocfilehash: 94e778bfa576354f3217431bac6a58558d624d11
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangeapply-method-publisher"></a>Метод ShapeRange.Apply (издатель)

Применяет форматирование, скопированные из другой фигуры или фигур с помощью метода **[раскладки](shaperange-pickup-method-publisher.md)** в диапазоне.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Применение**

 переменная _expression_A, представляющий объект **ShapeRange** .


### <a name="return-value"></a>Возвращаемое значение

Значение Nothing


## <a name="remarks"></a>Заметки

Если метод **раскладки** сначала не используется для копирования форматирования другую фигуру, возникает ошибка.


## <a name="example"></a>Пример

В следующем примере копируется форматирование из первой фигуры active публикации для второй фигуры active публикации.


```vb
With ActiveDocument.Pages(1) 
 .Shapes(1).PickUp 
 .Shapes(2).Apply 
End With 

```


