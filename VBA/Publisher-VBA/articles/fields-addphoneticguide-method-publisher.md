---
title: "Метод Fields.AddPhoneticGuide (издатель)"
keywords: vbapb10.chm6029320
f1_keywords: vbapb10.chm6029320
ms.prod: publisher
api_name: Publisher.Fields.AddPhoneticGuide
ms.assetid: 9b64e505-3aa7-040f-f791-f2dbeaf6860e
ms.date: 06/08/2017
ms.openlocfilehash: c31c3ebe412bd2287fe562eb8e3c3903789f0e22
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fieldsaddphoneticguide-method-publisher"></a>Метод Fields.AddPhoneticGuide (издатель)

Возвращает объект **[поля](field-object-publisher.md)** , представляющий фонетическое текст, добавляемый в указанный диапазон.


## <a name="syntax"></a>Синтаксис

 _выражение_. **AddPhoneticGuide** ( **_Диапазон_**, **_текст_**, **_Выравнивание_**, **_вызовет_**, **_FontName_**, **_FontSize_**)

 переменная _expression_A, представляющий объект **поля** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Range|Обязательное свойство.| **TextRange**|Текст в публикации, по которому фонетическое текст отображается.|
|Text|Обязательное свойство.| **String**|Фонетическое текст для добавления.|
|Выравнивание|Необязательный| **PbPhoneticGuideAlignmentType**|Выравнивание фонетическое добавлен текст.|
|Чтобы увеличить|Необязательный| **Variant**|Расстояние (в точках) в верхней части текста в указанном диапазоне в верхнюю часть фонетическое текста. Если значение не указано, Microsoft Publisher автоматически устанавливает фонетическое текст на оптимальное расстояние над указанным диапазоном.|
|FontName|Необязательный| **String**|Имя шрифта, используемого для фонетическое текста. Если значение не указано, используется тот же шрифт как текст в указанном диапазоне.|
|FontSize|Необязательный| **Variant**|Размер шрифта для фонетическое текста. Значение по умолчанию — 10 точек.|

### <a name="return-value"></a>Возвращаемое значение

Поле


## <a name="remarks"></a>Заметки

Параметр Выравнивание может иметь одно из **PbPhoneticGuideAlignmentType** константы объявляются в библиотеке типов Microsoft Publisher и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **pbPhoneticGuideAlignmentCenter**|Центры фонетическое текст в указанном диапазоне.|
| **pbPhoneticGuideAlignmentDefault**|Центры фонетическое текст в указанном диапазоне. По умолчанию.|
| **pbPhoneticGuideAlignmentLeft**| По левому краю фонетическое текст в указанном диапазоне.|
| **pbPhoneticGuideAlignmentOneTwoOne**|Настраивает внутри и вне интервал фонетическое текста в 1:2:1 отношение.|
| **pbPhoneticGuideAlignmentRight**|Правому краю фонетическое текст в указанном диапазоне.|
| **pbPhoneticGuideAlignmentZeroOneZero**|Настраивает внутри и вне интервал фонетическое текста в 0: соотношение 1:0.|

## <a name="example"></a>Пример

В этом примере добавляется фонетическое руководство для выбранного фразу «очень хорошо».


```vb
Sub PhoneticGuide() 
 Selection.TextRange.Fields.AddPhoneticGuide _ 
 Range:=Selection.TextRange, Text:="ver-E nIs", _ 
 Alignment:=pbPhoneticGuideAlignmentCenter, _ 
 Raise:=11, FontSize:=7 
End Sub
```


