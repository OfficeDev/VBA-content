---
title: "Метод TextRange.MoveStart (издатель)"
keywords: vbapb10.chm5308423
f1_keywords: vbapb10.chm5308423
ms.prod: publisher
api_name: Publisher.TextRange.MoveStart
ms.assetid: 5a9c480b-3cb7-0fd8-59c0-e2f93a925164
ms.date: 06/08/2017
ms.openlocfilehash: 0f9f52dbe7ebe8d6e77a515ad896f2c04902662e
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textrangemovestart-method-publisher"></a>Метод TextRange.MoveStart (издатель)

Перемещает начальное положение указанного диапазона. Этот метод возвращает значение типа **Long** , указывающее количество единиц, на которое начальное положение или диапазон или выделить фрагмент фактически перемещен или возвращает нуль (0), если не удалось выполнить перемещение.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Методов MoveStart** ( **_Единицы_**, **_размер_**)

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


