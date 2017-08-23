---
title: "Метод Table.ApplyAutoFormat (издатель)"
keywords: vbapb10.chm4784137
f1_keywords: vbapb10.chm4784137
ms.prod: publisher
api_name: Publisher.Table.ApplyAutoFormat
ms.assetid: f792a5f3-0d1c-06de-a030-7a588ca372d2
ms.date: 06/08/2017
ms.openlocfilehash: c3ed35e66edaef2e3709ec3de8abdc84cd39e068
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="tableapplyautoformat-method-publisher"></a>Метод Table.ApplyAutoFormat (издатель)

Область применения автоматического форматирования к указанной таблице встроенных таблицы.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ApplyAutoFormat** ( **_Автоформат_**, **_TextFormatting_**, **_TextAlignment_**, **_заполните поля_**, **_границы_**)

 переменная _expression_A, представляет собой объект- **таблицы** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Автоформат|Обязательное свойство.| **PbTableAutoFormatType**|Тип автоматического форматирования для применения к указанной таблицы.|
|TextFormatting|Необязательный| **Boolean**| **Значение true,** Чтобы применить форматирование текста в таблице шрифта. Значение по умолчанию — **True**.|
|TextAlignment|Необязательный| **Boolean**| **Значение true,** Чтобы применить выравнивание текста в таблице. Значение по умолчанию — **True**.|
|Заполните поля|Необязательный| **Boolean**| **Значение true,** Чтобы применить форматирование ячеек в таблице заливку. Значение по умолчанию — **True**.|
|Границы|Необязательный| **Boolean**| **Значение true,** Чтобы использовать границы для ячеек в таблице. Значение по умолчанию — **True**.|

## <a name="remarks"></a>Заметки

Параметр Автоформат может иметь одно из **[PbTableAutoFormatType](pbtableautoformattype-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.


## <a name="example"></a>Пример

В этом примере применяется фразу checkbook register автоматическое форматирование, заливки и границы для указанной таблицы.


```vb
Sub ApplyAutomaticTableFormatting() 
 ActiveDocument.Pages(1).Shapes(1).Table.ApplyAutoFormat _ 
 AutoFormat:=pbTableAutoFormatCheckbookRegister, _ 
 Borders:=False 
End Sub
```


