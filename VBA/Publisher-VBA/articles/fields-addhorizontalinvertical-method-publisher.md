---
title: "Метод Fields.AddHorizontalInVertical (издатель)"
keywords: vbapb10.chm6029319
f1_keywords: vbapb10.chm6029319
ms.prod: publisher
api_name: Publisher.Fields.AddHorizontalInVertical
ms.assetid: 4b451a24-0d79-70d4-4910-2725f1ed0297
ms.date: 06/08/2017
ms.openlocfilehash: d107f01c2b7cce41dbc038a292fbda3e4ce7d8e6
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fieldsaddhorizontalinvertical-method-publisher"></a>Метод Fields.AddHorizontalInVertical (издатель)

Вставляет горизонтальный текст в поток вертикальной и возвращает новый текст горизонтальной как объект **поля** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **AddHorizontalInVertical** ( **_Диапазон_**, **_текст_**)

 переменная _expression_A, представляющий объект **поля** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Range|Обязательное свойство.| **TextRange**|Диапазон текста для вставки горизонтальный текст.|
|Text|Обязательное свойство.| **String**|Текст, вставляемый по горизонтали.|

### <a name="return-value"></a>Возвращаемое значение

Поле


## <a name="example"></a>Пример

В этом примере по горизонтали вставляет текст «горизонтальной тест» после существующего вертикальный текст в форму одно на странице один из активных публикации.


```vb
Dim rngTemp As TextRange 
Dim fldTemp As Field 
 
With ActiveDocument.Pages(1).Shapes(1) 
 Set rngTemp = .TextFrame.TextRange.InsertAfter("") 
 
 Set fldTemp = .TextFrame.TextRange.Fields _ 
 .AddHorizontalInVertical(Range:=rngTemp, Text:="horizontal test") 
End With
```


