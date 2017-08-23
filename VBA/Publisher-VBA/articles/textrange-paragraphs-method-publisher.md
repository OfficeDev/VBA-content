---
title: "Метод TextRange.Paragraphs (издатель)"
keywords: vbapb10.chm5308454
f1_keywords: vbapb10.chm5308454
ms.prod: publisher
api_name: Publisher.TextRange.Paragraphs
ms.assetid: 895c32cf-cdbe-74b0-ab47-6ae63d1bdea0
ms.date: 06/08/2017
ms.openlocfilehash: 7871fd8e90e028f8cbf13febc10be655ecff07c8
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textrangeparagraphs-method-publisher"></a>Метод TextRange.Paragraphs (издатель)

Возвращает объект **[TextRange](textrange-object-publisher.md)** , представляющий указанного абзацев.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Абзацы** ( **_Запуск_**, **_Длина_**)

 переменная _expression_A, представляющий объект **TextRange** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Начало|Обязательное свойство.| **Длинный**|Первый абзац возвращаемого диапазона.|
|Length|Необязательный| **Длинный**|Число абзацев которого требуется получить. Значение по умолчанию — 1.|

### <a name="return-value"></a>Возвращаемое значение

TextRange


## <a name="example"></a>Пример

Если **_Длина_** опущен, возвращенный диапазон содержит один абзац.



Если **_Длина_** больше, чем количество абзацев из указанного начального абзаца по завершению теста, возвращенный диапазон содержит все абзацы.

В этом примере преобразует в формат отступы первой строки выделенного абзаца.




```vb
Sub FormatCurrentParagraph() 
 Selection.TextRange.Paragraphs(Start:=1).ParagraphFormat _ 
 .FirstLineIndent = InchesToPoints(0.5) 
End Sub
```


