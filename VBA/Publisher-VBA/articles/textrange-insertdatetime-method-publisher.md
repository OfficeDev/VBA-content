---
title: "Метод TextRange.InsertDateTime (издатель)"
keywords: vbapb10.chm5308453
f1_keywords: vbapb10.chm5308453
ms.prod: publisher
api_name: Publisher.TextRange.InsertDateTime
ms.assetid: 1d02471a-f22b-7dad-bcbb-40af3a04d198
ms.date: 06/08/2017
ms.openlocfilehash: 8a9a0b177e0034dc7ad14afd581f05eeecf4ce35
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textrangeinsertdatetime-method-publisher"></a>Метод TextRange.InsertDateTime (издатель)

Возвращает объект **[TextRange](textrange-object-publisher.md)** , представляющий дату и время в диапазоне указанный текст.


## <a name="syntax"></a>Синтаксис

 _выражение_. **InsertDateTime** ( **_Формат_**, **_InsertAsField_**, **_InsertAsFullWidth_**, **_язык_**, **_Календарь_**)

 переменная _expression_A, представляющий объект **TextRange** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Формат|Обязательное свойство.| **PbDateTimeFormat**|Формат даты и времени.|
|InsertAsField|Необязательный| **Boolean**| **Значение true** для Microsoft Publisher для обновления даты и времени при каждом открытии публикации. Значение по умолчанию — **False**.|
|InsertAsFullWidth|Необязательный| **Boolean**| **Значение true** для вставки указанных данных в виде двухбайтовых разрядов. Этот аргумент могут быть недоступны, в зависимости от языка Английский (США, например), выбранных или установленных. Значение по умолчанию — **False**.|
|Language|Необязательный| **MsoLanguageID**|Язык, на котором для отображения значений даты и времени.|
|Календарь|Необязательный| **PbCalendarType**|Тип календаря, используемый для отображения значений даты и времени.|

### <a name="return-value"></a>Возвращаемое значение

TextRange


## <a name="remarks"></a>Заметки

Параметр Format может иметь одно из **[PbDateTimeFormat](pbdatetimeformat-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.

Параметр Language может иметь одно из ** [MsoLanguageID](http://msdn.microsoft.com/library/65ea40f0-9a09-3d76-1519-4acddcc5f367%28Office.15%29.aspx)** объявленные константы в библиотеке типов, Microsoft Office.

Параметр календарь может иметь одно из **[PbCalendarType](pbcalendartype-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher. Значение по умолчанию — **pbCalendarTypeWestern**.


## <a name="example"></a>Пример

В этом примере Вставка поля для текущей даты в позиции курсора.


```vb
Sub InsertDateField() 
 Selection.TextRange.InsertDateTime Format:=pbDateLong, InsertAsField:=True 
End Sub
```


