---
title: "Свойство ReaderSpread.Pages (издатель)"
keywords: vbapb10.chm524293
f1_keywords: vbapb10.chm524293
ms.prod: publisher
api_name: Publisher.ReaderSpread.Pages
ms.assetid: 181c37b2-ed3f-826a-5718-ae6aff120eb3
ms.date: 06/08/2017
ms.openlocfilehash: 00d0f9c2fc7508f1af937d92ad9252e905f52c17
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="readerspreadpages-property-publisher"></a>Свойство ReaderSpread.Pages (издатель)

Возвращает объект **[Page](page-object-publisher.md)** , представляющий одну из страниц, составляющих указанного ширина чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Страницы** ( **_Индекс_**)

 переменная _expression_A, представляет собой объект- **ReaderSpread** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Индекс|Обязательное свойство.| **Длинный**|Для возврата распространение страницы из устройства чтения. Может быть 1 или 2.|

## <a name="remarks"></a>Заметки

Средство чтения распространения будет состоять из только для одного или двух страниц, поэтому поддерживаются следующие допустимые значения для аргумента **Index** 1 или 2.


## <a name="example"></a>Пример

В следующем примере проверяется распространения чтения странице пятый в активной публикации на предмет наличия более одной страницы. Если это так, в примере сообщается номер второй страницы в распространении.


```vb
Dim pageTemp As Page 
 
With ActiveDocument.Pages(5).ReaderSpread 
 If .PageCount > 1 Then 
 Set pageTemp = .Pages(Index:=2) 
 MsgBox "The page number of the second page " _ 
 &; "in the spread is " &; pageTemp.PageNumber 
 Else 
 MsgBox "The spread has only one page." 
 End If 
End With
```


