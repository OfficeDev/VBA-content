---
title: "Свойство ShapeRange.BlackWhiteMode (издатель)"
keywords: vbapb10.chm2293872
f1_keywords: vbapb10.chm2293872
ms.prod: publisher
api_name: Publisher.ShapeRange.BlackWhiteMode
ms.assetid: c85babbd-f05d-c3e1-3265-c08888eaf212
ms.date: 06/08/2017
ms.openlocfilehash: 9b08901bc3f1adb78e2227f819435e7d15302d15
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangeblackwhitemode-property-publisher"></a>Свойство ShapeRange.BlackWhiteMode (издатель)

Возвращает или задает константой **MsoBlackWhiteMode**, указывающее, как указанные форму или диапазона фигуры отображается при просмотре публикации в черно-белом режиме. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **BlackWhiteMode**

 переменная _expression_A, представляющий объект **ShapeRange** .


## <a name="remarks"></a>Заметки

Значение свойства **BlackWhiteMode** может иметь одно из ** [MsoBlackWhiteMode](http://msdn.microsoft.com/library/2b4d7e22-1277-9f5c-ba52-a37e113477c1%28Office.15%29.aspx)** объявленные константы в библиотеке типов, Microsoft Office.


## <a name="example"></a>Пример

В этом примере задается первую фигуру в active публикации, которая будет отображаться в черно-белом режиме. При просмотре публикации в черно-белом режиме фигуры будет отображаться черные, независимо от того, какой цвет в режиме цвета.


```vb
ActiveDocument.Pages(1).Shapes(1).BlackWhiteMode = msoBlackWhiteBlack
```


