---
title: "Свойство Font.UseDiacriticColor (издатель)"
keywords: vbapb10.chm5374002
f1_keywords: vbapb10.chm5374002
ms.prod: publisher
api_name: Publisher.Font.UseDiacriticColor
ms.assetid: 368d3599-b0b0-1790-0ce0-13f1936bccb0
ms.date: 06/08/2017
ms.openlocfilehash: f6331a37c8a1bf35e39945d9e1b6b68fd5e7c874
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fontusediacriticcolor-property-publisher"></a>Свойство Font.UseDiacriticColor (издатель)

Возвращает или задает константу **MsoTriState** , указывающее, можно ли установить цвет диакритические знаки в диапазоне указанный текст. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **UseDiacriticColor**

 переменная _expression_A, представляющий объект **шрифта** .


### <a name="return-value"></a>Возвращаемое значение

MsoTriState


## <a name="remarks"></a>Заметки

Значение свойства **UseDiacriticColor** может иметь одно из **MsoTriState** константы объявляются в библиотеке типов Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**|Цвет диакритические знаки нельзя задать в диапазоне указанный текст.|
| **msoTriStateMixed**|Возвращает значение, указывающее, сочетание **msoTrue** и **msoFalse** для диапазона указанной фигуры.|
| **msoTriStateToggle**|Задайте значение, могут переключаться между **msoTrue** и **msoFalse**.|
| **msoTrue**|Позволяет указать цвет диакритические знаки в диапазоне указанный текст.|

## <a name="example"></a>Пример

В этом примере проверить текст в первой статьи публикации для состояния свойства **UseDiacriticColor** . Если это **msoTrue**синий задано значение свойства **DiacriticColor** . В противном случае отображается окно сообщения.


```vb
Sub UseDiaColor() 
 
 Dim fntDC As Font 
 
 Set fntDC = Application.ActiveDocument.Stories(1).TextRange.Font 
 If fntDC.UseDiacriticColor = msoTrue Then 
 fntDC.DiacriticColor.RGB = RGB(Red:=0, Green:=0, Blue:=255) 
 Else 
 MsgBox "The UseDiacriticColor property is set to False" 
 End If 
 
End Sub
```


