---
title: "Свойство Shape.VerticalFlip (издатель)"
keywords: vbapb10.chm2228308
f1_keywords: vbapb10.chm2228308
ms.prod: publisher
api_name: Publisher.Shape.VerticalFlip
ms.assetid: b3c7492f-08ee-8fad-102a-8e2a2f69b969
ms.date: 06/08/2017
ms.openlocfilehash: b77847f31d26d7eccc15d1111e168d6b3c3c3a81
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapeverticalflip-property-publisher"></a>Свойство Shape.VerticalFlip (издатель)

Возвращает **msoTrue** , если указанный фигуры отразилось вокруг вертикальной оси. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **VerticalFlip**

 переменная _expression_A, представляющий объект **фигуры** .


## <a name="remarks"></a>Заметки

Значение свойства может быть одной из констант **MsoTriState** объявлена в библиотеке типов, Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**|Фигура не отразилось вокруг вертикальной оси.|
| **msoTriStateMixed**|Указывает сочетание **msoTrue** и **msoFalse** для диапазона указанной фигуры.|
| **msoTrue**|Фигура отразилось вокруг вертикальной оси.|

## <a name="example"></a>Пример

В этом примере восстанавливает исходное состояние каждой фигуры на активной публикации, если его отразилось по горизонтали или по вертикали.


```vb
Sub Flipper() 
 
 Dim shpBall As Shape 
 
 For Each shpBall In ActiveDocument.MasterPages.Item(1).Shapes 
 If shpBall.HorizontalFlip = msoTrue Then shpBall.Flip msoFlipHorizontal 
 If shpBall.VerticalFlip = msoTrue Then shpBall.Flip msoFlipVertical 
 Next 
 
End Sub
```


