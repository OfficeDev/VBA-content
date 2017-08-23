---
title: "Метод TextRange.MoveEnd (издатель)"
keywords: vbapb10.chm5308424
f1_keywords: vbapb10.chm5308424
ms.prod: publisher
api_name: Publisher.TextRange.MoveEnd
ms.assetid: 4fe27375-34e2-2ecc-33c8-a07230012b13
ms.date: 06/08/2017
ms.openlocfilehash: d6a4d3bd57c5b5bac43819a9de7de7d24de2f9fd
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textrangemoveend-method-publisher"></a>Метод TextRange.MoveEnd (издатель)

Перемещает конечного диапазона символов. Этот метод возвращает **Long** , представляющее номер единиц измерения диапазон или выбора фактически перемещается или возвращает нуль (0), если не удалось выполнить перемещение.


## <a name="syntax"></a>Синтаксис

 _выражение_. **MoveEnd** ( **_Единицы_**, **_размер_**)

 переменная _expression_A, представляющий объект **TextRange** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Подразделения|Обязательное свойство.| **PbTextUnit**|Подразделения, с помощью которого является перемещены свернутые диапазон или выделить фрагмент.|
|Размер|Обязательное свойство.| **Длинный**|Число единиц измерения для перемещения. Если этот номер является положительным, положение конечного знака перемещается вперед в документе. Если этот номер является отрицательным, конца перемещается назад. Если положение конечного положением начального знака, диапазон сворачивается, а оба символа положения перемещаются одновременно.|

### <a name="return-value"></a>Возвращаемое значение

Длинный


## <a name="remarks"></a>Заметки

Параметр устройства может быть одной из констант **PbTextUnit** объявлена в библиотеке типов, Microsoft Publisher и показаны в следующей таблице.



| **pbTextUnitCell**|| **pbTextUnitCharacter**|| **pbTextUnitCharFormat**|| **pbTextUnitCodePoint**|| **pbTextUnitColumn**|| **pbTextUnitLine**|| **pbTextUnitObject**|| **pbTextUnitParaFormat**|| **pbTextUnitParagraph**|| **pbTextUnitRow**|| **pbTextUnitScreen**|| **pbTextUnitSection**|| **pbTextUnitSentence**|| **pbTextUnitStory**|| **pbTextUnitTable**|| **pbTextUnitWindow**|| **pbTextUnitWord**|

## <a name="example"></a>Пример

В этом примере задает диапазон текста, перемещает диапазон начальной и конечной позиций символов и форматирует шрифта для диапазона.


```vb
Sub MoveStartEnd() 
 Dim rngText As TextRange 
 
 Set rngText = ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .TextRange.Paragraphs(Start:=3, Length:=1) 
 
 With rngText 
 .MoveStart Unit:=pbTextUnitLine, Size:=-2 
 .MoveEnd Unit:=pbTextUnitLine, Size:=1 
 With .Font 
 .Bold = msoTrue 
 .Size = 15 
 End With 
 End With 
 
End Sub
```


