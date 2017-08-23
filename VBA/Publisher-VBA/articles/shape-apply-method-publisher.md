---
title: "Метод Shape.Apply (издатель)"
keywords: vbapb10.chm2228240
f1_keywords: vbapb10.chm2228240
ms.prod: publisher
api_name: Publisher.Shape.Apply
ms.assetid: 711c72b6-3618-be0b-fb72-9f68fdbcc4a8
ms.date: 06/08/2017
ms.openlocfilehash: b5099e18a296982124e036fc552dc19560e824c9
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapeapply-method-publisher"></a>Метод Shape.Apply (издатель)

Применяет форматирование, скопированные из другой фигуры или фигур с помощью метода **[раскладки](shape-pickup-method-publisher.md)** в диапазоне.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Применение**

 переменная _expression_A, представляющий объект **фигуры** .


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


