---
title: "Метод ParagraphFormat.SetLineSpacing (издатель)"
keywords: vbapb10.chm5439511
f1_keywords: vbapb10.chm5439511
ms.prod: publisher
api_name: Publisher.ParagraphFormat.SetLineSpacing
ms.assetid: 32e5b233-8415-2373-7423-18b66df3a5ea
ms.date: 06/08/2017
ms.openlocfilehash: fc41da03b89bb914ff8daadbc3278d6ba7e336f6
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="paragraphformatsetlinespacing-method-publisher"></a>Метод ParagraphFormat.SetLineSpacing (издатель)

Форматирует междустрочным интервалом указанного абзацев.


## <a name="syntax"></a>Синтаксис

 _выражение_. **SetLineSpacing** ( **_Правило_**, **_интервал_**)

 переменная _expression_A, представляет собой объект- **ParagraphFormat** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Rule|Обязательное свойство.| **PbLineSpacingRule**|Междустрочным интервалом для указанного абзацев.|
|Интервал|Необязательный| **Variant**|Интервал (в пунктах) для указанного абзацев.|

## <a name="remarks"></a>Заметки

Параметр правила может быть одной из констант **PbLineSpacingRule** объявлена в библиотеке типов, Microsoft Publisher и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **pbLineSpacing1pt5**|Задает интервал для указанного абзацев-и a две строки.|
| **pbLineSpacingDouble**| Двойные пробелы содержатся указанного абзацев.|
| **pbLineSpacingExactly**| Задает междустрочным интервалом точно значению, указанному в аргументе интервал даже в том случае, если размер шрифта используется внутри абзаца.|
| **pbLineSpacingMixed**| Возвращаемое значение для свойства **[LineSpacing](paragraphformat-linespacing-property-publisher.md)** , которое указывает, что междустрочным интервалом — это сочетание значений для указанного абзацев.|
| **pbLineSpacingMultiple**|Установка интервала строки значению, указанному в аргументе интервал.|
| **pbLineSpacingSingle**|Один пробелы указанного абзацев.|

## <a name="example"></a>Пример

В этом примере задается междустрочным интервалом в double.


```vb
Sub SetLineSpacingForSelection() 
 Selection.TextRange.ParagraphFormat.SetLineSpacing _ 
 Rule:=pbLineSpacingDouble, Spacing:=12 
End Sub
```


