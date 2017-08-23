---
title: "Свойство TextFrame.Overflowing (издатель)"
keywords: vbapb10.chm3866649
f1_keywords: vbapb10.chm3866649
ms.prod: publisher
api_name: Publisher.TextFrame.Overflowing
ms.assetid: 5a0f053b-519a-1637-0d73-992c56cdd7f0
ms.date: 06/08/2017
ms.openlocfilehash: 76d639030e5ba6cf92f610128ae41d6720cbfb1f
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textframeoverflowing-property-publisher"></a>Свойство TextFrame.Overflowing (издатель)

Показывает, содержит ли рамки больше текста, чем помещается в текстовую рамку. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Переполнения**

 переменная _expression_A, представляющий объект **TextFrame** .


### <a name="return-value"></a>Возвращаемое значение

MsoTriState


## <a name="remarks"></a>Заметки

Значение свойства **Overflowing** может иметь одно из **MsoTriState** константы объявляются в библиотеке типов Microsoft Office и показаны в следующей таблице.



|**Константы**|
|:-----|
| **msoFalse**|
| **msoTrue**|

## <a name="example"></a>Пример

В этом примере увеличивает высота кадра выделенного текста, если он содержит избыточные текста.


```vb
Sub IncreaseTextBoxHeight() 
 With Selection.ShapeRange.TextFrame 
 If .Overflowing = msoTrue Then 
 Do 
 .Parent.Height = .Parent.Height + 36 
 Loop Until .Overflowing = msoFalse 
 End If 
 End With 
End Sub
```


