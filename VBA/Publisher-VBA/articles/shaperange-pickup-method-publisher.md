---
title: "Метод ShapeRange.PickUp (издатель)"
keywords: vbapb10.chm2293795
f1_keywords: vbapb10.chm2293795
ms.prod: publisher
api_name: Publisher.ShapeRange.PickUp
ms.assetid: ebd62b6e-807a-821c-d8ea-ed9be289c433
ms.date: 06/08/2017
ms.openlocfilehash: 1f16e34ac20c6b7b99df552bd37ee72ba0bca024
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangepickup-method-publisher"></a>Метод ShapeRange.PickUp (издатель)

Копирование форматирования фигуры или диапазона фигуры, чтобы его можно скопировать в другую фигуру или фигур с помощью метода **[Применить](shaperange-apply-method-publisher.md)** диапазона.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Раскладки**

 переменная _expression_A, представляющий объект **ShapeRange** .


## <a name="remarks"></a>Заметки

Необходимо использовать метод **раскладки** для копирования форматирования фигуры или диапазона фигуры перед использованием методу **Apply** ; в противном случае возникает ошибка.


## <a name="example"></a>Пример

В следующем примере копируется форматирование из первой фигуры active публикации для второй фигуры active публикации.


```vb
With ActiveDocument.Pages(1) 
 .Shapes(1).PickUp 
 .Shapes(2).Apply 
End With 

```


