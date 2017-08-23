---
title: "Метод Shape.PickUp (издатель)"
keywords: vbapb10.chm2228259
f1_keywords: vbapb10.chm2228259
ms.prod: publisher
api_name: Publisher.Shape.PickUp
ms.assetid: 12b59235-db2d-b451-de8e-9e8df6bfeb1c
ms.date: 06/08/2017
ms.openlocfilehash: da7f7a6c6afcc4c3ea2fc3a3699378ffa875557b
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapepickup-method-publisher"></a>Метод Shape.PickUp (издатель)

Копирование форматирования фигуры или диапазона фигуры, чтобы его можно скопировать в другую фигуру или фигур с помощью метода **[Применить](shaperange-apply-method-publisher.md)** диапазона.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Раскладки**

 переменная _expression_A, представляющий объект **фигуры** .


### <a name="return-value"></a>Возвращаемое значение

Значение Nothing


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


