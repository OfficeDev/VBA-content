---
title: "Метод TextRange.Move (издатель)"
keywords: vbapb10.chm5308422
f1_keywords: vbapb10.chm5308422
ms.prod: publisher
api_name: Publisher.TextRange.Move
ms.assetid: a51b4153-2ac5-2293-d2a0-d4a3786268d7
ms.date: 06/08/2017
ms.openlocfilehash: aa5c5f5fba3601c92f671583d810cdd795e6691a
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textrangemove-method-publisher"></a>Метод TextRange.Move (издатель)

Сворачивает его начала или окончания позиции указанного диапазона, а затем перемещает свернутые объекта указанное число единиц измерения. Этот метод возвращает значение типа **Long** , представляющее количество единиц, на которое объект действительно был перемещен или возвращает нуль (0), если не удалось выполнить перемещение.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Перемещение** ( **_Единицы_**, **_размер_**)

 переменная _expression_A, представляющий объект **TextRange** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Подразделения|Обязательное свойство.| **PbTextUnit**|Подразделения, с помощью которого является перемещены свернутые диапазон или выделить фрагмент.|
|Размер|Обязательное свойство.| **Длинный**|Число единиц измерения, по которым указанный диапазон или выделить фрагмент не переместить. Если **размер** — это положительное число, объект свернуты в конец положение и переместить вперед в документе на указанное число единиц измерения. Если **размер** отрицательное значение, объект свернуты в положение начала и переместить назад указанное число единиц измерения. Направление свернуть также можно управлять с помощью метода **Свернуть** перед использованием метода **Move** .|

### <a name="return-value"></a>Возвращаемое значение

Длинный


## <a name="remarks"></a>Заметки

Параметр устройства может быть одной из констант **PbTextUnit** объявлена в библиотеке типов, Microsoft Publisher и показаны в следующей таблице.



| **pbTextUnitCell**|| **pbTextUnitCharacter**|| **pbTextUnitCharFormat**|| **pbTextUnitCodePoint**|| **pbTextUnitColumn**|| **pbTextUnitLine**|| **pbTextUnitObject**|| **pbTextUnitParaFormat**|| **pbTextUnitParagraph**|| **pbTextUnitRow**|| **pbTextUnitScreen**|| **pbTextUnitSection**|| **pbTextUnitSentence**|| **pbTextUnitStory**|| **pbTextUnitTable**|| **pbTextUnitWindow**|| **pbTextUnitWord**|

## <a name="example"></a>Пример

В этом примере сворачивает указанного диапазона и вставка нового предложения в начало диапазона.


```vb
Sub MoveText() 
 Dim rngText As TextRange 
 Set rngText = ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .TextRange.Words(Start:=1, Length:=5) 
 With rngText 
 .Move Unit:=pbTextUnitParagraph, Size:=-1 
 .Text = "This adds new text to the beginning of the range. " 
 End With 
End Sub
```


