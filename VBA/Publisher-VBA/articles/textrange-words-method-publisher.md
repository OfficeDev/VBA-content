---
title: "Метод TextRange.Words (издатель)"
keywords: vbapb10.chm5308456
f1_keywords: vbapb10.chm5308456
ms.prod: publisher
api_name: Publisher.TextRange.Words
ms.assetid: df812db2-98ca-848b-7922-6905cb71124c
ms.date: 06/08/2017
ms.openlocfilehash: a7616bd89fe8decd1f231acfeab9fd54cb380fc3
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textrangewords-method-publisher"></a>Метод TextRange.Words (издатель)

Возвращает объект **[TextRange](textrange-object-publisher.md)** , представляющий указанного подмножества слова.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Слова** ( **_Запуск_**, **_Длина_**)

 переменная _expression_A, представляющий объект **TextRange** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Начало|Обязательное свойство.| **Длинный**|Первое слово возвращаемого диапазона.|
|Length|Необязательный| **Длинный**|Количество слов должно быть возвращено. Значение по умолчанию — 1.|

### <a name="return-value"></a>Возвращаемое значение

TextRange


## <a name="remarks"></a>Заметки

Если **_Длина_** опущен, возвращенный диапазон содержит одно слово.

Если **_запустить_** больше, чем количество слов в указанный текст, возвращенный диапазон начинается с последнего слова в указанном диапазоне.

Если **_Длина_** больше, чем количество слов из указанного начального word в конец текста, возвращенный диапазон содержит эти слова.


## <a name="example"></a>Пример

В этом примере форматов как полужирный шрифт секунды в-третьих, и четвертый слова, набранные фигуры два по одному active публикации.


```vb
Application.ActiveDocument.Pages(1).Shapes(2) _ 
 .TextFrame.TextRange.Words(Start:=2, Length:=3) _ 
 .Font.Bold = True
```


