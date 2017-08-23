---
title: "Свойство ShapeRange.HorizontalFlip (издатель)"
keywords: vbapb10.chm2293824
f1_keywords: vbapb10.chm2293824
ms.prod: publisher
api_name: Publisher.ShapeRange.HorizontalFlip
ms.assetid: c0dd2f4a-0baf-3720-113a-b929193f2b1d
ms.date: 06/08/2017
ms.openlocfilehash: 48745b9759c9759f5948b6b93abeb116415a41c5
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangehorizontalflip-property-publisher"></a>Свойство ShapeRange.HorizontalFlip (издатель)

Указывает, является ли указанный фигуры отразилось относительно его горизонтальной оси. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **HorizontalFlip**

 переменная _expression_A, представляющий объект **ShapeRange** .


## <a name="remarks"></a>Заметки

Значение свойства **HorizontalFlip** может иметь одно из **MsoTriState** константы объявляются в библиотеке типов Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**|Фигура не отразилось относительно его горизонтальной оси.|
| **msoTriStateMixed**|Указывает сочетание **msoTrue** и **msoFalse** для диапазона указанной фигуры.|
| **msoTrue**|Фигура отразилось относительно его горизонтальной оси.|

## <a name="example"></a>Пример

В этом примере восстанавливает исходное состояние каждой фигуры на активной публикации, если его отразилось по горизонтали или по вертикали.


```vb
Sub Flipper() 
 
 Dim shpS As Shape 
 
 For Each shpS In ActiveDocument.MasterPages.Item(1).Shapes 
 If shpS.HorizontalFlip = msoTrue Then shpS.Flip msoFlipHorizontal 
 If shpS.VerticalFlip = msoTrue Then shpS.Flip msoFlipVertical 
 Next 
 
End Sub
```


