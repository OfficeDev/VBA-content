---
title: "Свойство ShapeRange.VerticalFlip (издатель)"
keywords: vbapb10.chm2293844
f1_keywords: vbapb10.chm2293844
ms.prod: publisher
api_name: Publisher.ShapeRange.VerticalFlip
ms.assetid: cc3ab3ec-71f6-49fc-0141-505054d6abbb
ms.date: 06/08/2017
ms.openlocfilehash: dd509dd5817498114ceeeeb48baff3e5374f1b7a
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangeverticalflip-property-publisher"></a>Свойство ShapeRange.VerticalFlip (издатель)

Возвращает **msoTrue** , если указанный фигуры отразилось вокруг вертикальной оси. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **VerticalFlip**

 переменная _expression_A, представляющий объект **ShapeRange** .


## <a name="remarks"></a>Заметки

Значение свойства **VerticalFlip** может иметь одно из **MsoTriState** константы объявляются в библиотеке типов Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**|Фигура не отразилось вокруг вертикальной оси.|
| **msoTriStateMixed**|Возвращает значение, указывающее, сочетание **msoTrue** и **msoFalse** для диапазона указанной фигуры.|
| **msoTriStateToggle**|Задайте значение, могут переключаться между **msoTrue** и **msoFalse**.|
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


