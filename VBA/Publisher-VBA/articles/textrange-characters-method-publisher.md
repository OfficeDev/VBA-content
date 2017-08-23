---
title: "Метод TextRange.Characters (издатель)"
keywords: vbapb10.chm5308425
f1_keywords: vbapb10.chm5308425
ms.prod: publisher
api_name: Publisher.TextRange.Characters
ms.assetid: e851767e-12b2-ad77-071b-9d27bbf0d637
ms.date: 06/08/2017
ms.openlocfilehash: 944d5f489175c4eda06f0ff86e25c58f56932f4f
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textrangecharacters-method-publisher"></a>Метод TextRange.Characters (издатель)

Возвращает объект **[TextRange](textrange-object-publisher.md)** , представляющий указанного подмножества текстовых символов.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Символы** ( **_Запуск_**, **_Длина_**)

 переменная _expression_A, представляющий объект **TextRange** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Начало|Обязательное свойство.| **Длинный**|Первый символ возвращаемого диапазона.|
|Length|Необязательный| **Длинный**|Число символов, которые будут возвращены. Значение по умолчанию — 1.|

### <a name="return-value"></a>Возвращаемое значение

TextRange


## <a name="remarks"></a>Заметки

Если **_запустить_** больше, чем количество символов в указанный текст, возвращенный диапазон начинается с последнего символа в указанном диапазоне.

Если **_Длина_** больше, чем количество символов из указанного начального знака в конец текста, возвращенный диапазон содержит все символы.


## <a name="example"></a>Пример

В этом примере задает текст для первой фигуры на первой странице в активном документе и затем задает шрифт первые два разряда 15 точек и полужирным шрифтом.


```vb
Sub CharRange() 
 Dim rngCharacters As TextRange 
 Set rngCharacters = Application.ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.InsertBefore(NewText:="Hello World.") 
 With rngCharacters.Characters(Start:=1, Length:=2).Font 
 .Size = 15 
 .Bold = msoTrue 
 End With 
End Sub
```


