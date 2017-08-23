---
title: "Метод TextRange.Lines (издатель)"
keywords: vbapb10.chm5308455
f1_keywords: vbapb10.chm5308455
ms.prod: publisher
api_name: Publisher.TextRange.Lines
ms.assetid: 56862090-b2ff-403b-d016-e37108d5ccc1
ms.date: 06/08/2017
ms.openlocfilehash: 5519e0f7f25e6ded73cceaba404d29ff63bf574a
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textrangelines-method-publisher"></a>Метод TextRange.Lines (издатель)

Возвращает объект **[TextRange](textrange-object-publisher.md)** , представляющий указанной строки.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Строки** ( **_Запуск_**, **_Длина_**)

 переменная _expression_A, представляющий объект **TextRange** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Начало|Обязательное свойство.| **Длинный**|Первая строка возвращаемого диапазона.|
|Length|Необязательный| **Длинный**|Число строк должно быть возвращено. Значение по умолчанию — 1.|

### <a name="return-value"></a>Возвращаемое значение

TextRange


## <a name="remarks"></a>Заметки

Если **_запустить_** больше, чем количество строк в указанный текст, возвращенный диапазон начинается с последней строки в указанном диапазоне.

Если **_Длина_** больше, чем количество строк из указанного начальную строку в конец текста, возвращенный диапазон содержит все строки.


## <a name="example"></a>Пример

В этом примере заменяет первые три строки первую фигуру на первой странице указанной строки.


```vb
Sub ReplaceLines() 
 Dim rngText As TextRange 
 Set rngText = ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.Lines(Start:=1, Length:=3) 
 
 rngText.Text = "This is replacement text." &; vbCrLf 
 
End Sub
```


