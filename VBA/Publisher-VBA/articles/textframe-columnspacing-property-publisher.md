---
title: "Свойство TextFrame.ColumnSpacing (издатель)"
keywords: vbapb10.chm3866633
f1_keywords: vbapb10.chm3866633
ms.prod: publisher
api_name: Publisher.TextFrame.ColumnSpacing
ms.assetid: 3b650d29-3716-e9b1-eaf0-92bdc0b77c5f
ms.date: 06/08/2017
ms.openlocfilehash: a4a1959558346a0ba0521f7528a314e8f9842aa2
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textframecolumnspacing-property-publisher"></a>Свойство TextFrame.ColumnSpacing (издатель)

Возвращает или задает **Variant** , который представляет дискового пространства между столбцами текста. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ColumnSpacing**

 переменная _expression_A, представляет собой объект- **TextFrame** .


### <a name="return-value"></a>Возвращаемое значение

Variant


## <a name="remarks"></a>Заметки

Интервал меры, начиная с конца текста до конца столбца и еще раз с самого начала столбца в начало текста. Таким образом при вводе суммы **ColumnSpacing** 0,5 дюйма, общее интервалы между столбцы — это один дюйм: 0,5 дюйма измерение в конце текста в конец столбца в один столбец, а также 0,5 дюйма измерение с самого начала столбца в начало текста в соседних столбцов.


## <a name="example"></a>Пример

В этом примере форматов первым текстовым полем в активной публикации с трех столбцов и всего 0,5 дюйма интервалы между столбцов.


```vb
Sub SetColumnsAndSpacing() 
 With ActiveDocument.Pages(1).Shapes(1).TextFrame 
 .Columns = 3 
 .ColumnSpacing = InchesToPoints(0.25) 
 End With 
End Sub
```


