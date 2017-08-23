---
title: "Метод TextRange.Expand (издатель)"
keywords: vbapb10.chm5308421
f1_keywords: vbapb10.chm5308421
ms.prod: publisher
api_name: Publisher.TextRange.Expand
ms.assetid: 66d8b1a3-5fc4-bed7-94d2-06be6203e1e9
ms.date: 06/08/2017
ms.openlocfilehash: 57ba322cea87c1fe0305b9740fdf948ab0d85b05
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textrangeexpand-method-publisher"></a>Метод TextRange.Expand (издатель)

При развертывании указанный диапазон или выделить фрагмент. Возвращает или задает типа **Long** , представляющее номер указанного единицы, добавляемого диапазон или выделить фрагмент.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Разверните узел** ( **_Единицы_**)

 переменная _expression_A, представляющий объект **TextRange** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Подразделения|Обязательное свойство.| **PbTextUnit**|Единица, на которое нужно расширить диапазон.|

### <a name="return-value"></a>Возвращаемое значение

Длинный


## <a name="remarks"></a>Заметки

Метод **развертывания** перемещает обе конечные точки диапазона, если это необходимо; Чтобы переместить только одну конечную точку диапазона, используйте метод **[методов MoveStart](textrange-movestart-method-publisher.md)** и **[MoveEnd](textrange-moveend-method-publisher.md)** .

Параметр устройства может иметь одно из **[PbTextUnit](pbtextunit-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.


## <a name="example"></a>Пример

В этом примере создается диапазон, который ссылается на первое слово в первую фигуру active публикации, шрифта, используемого для word, а затем его расширяет диапазон ссылок на всей первый абзац и форматирование шрифта для всей строки.


```vb
Sub ExpandRange() 
 Dim rngText As TextRange 
 
 Set rngText = ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.Words(Start:=1, Length:=1) 
 With rngText 
 With .Font 
 .Size = 20 
 .Italic = msoTrue 
 End With 
 .Expand Unit:=pbTextUnitLine 
 .Font.Bold = msoTrue 
 End With 
End Sub
```


