---
title: "Метод TextRange.InsertPageNumber (издатель)"
keywords: vbapb10.chm5308486
f1_keywords: vbapb10.chm5308486
ms.prod: publisher
api_name: Publisher.TextRange.InsertPageNumber
ms.assetid: f71d3b40-0263-93fa-d7e3-d815b90f71f7
ms.date: 06/08/2017
ms.openlocfilehash: b3a9e2f07320a86b1cb77e20c3ec45ffd875279c
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textrangeinsertpagenumber-method-publisher"></a>Метод TextRange.InsertPageNumber (издатель)

Возвращает объект **[TextRange](textrange-object-publisher.md)** , представляющий поля номера страницы в публикации.


## <a name="syntax"></a>Синтаксис

 _выражение_. **InsertPageNumber** ( **_Тип_**)

 переменная _expression_A, представляющий объект **TextRange** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Тип|Необязательный| **PbPageNumberType**|Задает, является ли номер страницы в текущий номер страницы или номер страницы следующий или предыдущий связанный текстового поля.|

### <a name="return-value"></a>Возвращаемое значение

TextRange


## <a name="remarks"></a>Заметки

Тип может иметь одно из следующих констант **PbPageNumberType** .



|**Константы**|**Описание**|
|:-----|:-----|
| **pbPageNumberCurrent**|По умолчанию.|
| **pbPageNumberNextInStory**|Вставляет номер следующей связанной надписи.|
| **pbPageNumberPreviousInStory**|Вставляет номер предыдущей связанной надписи.|

## <a name="example"></a>Пример

В этом примере Вставка поля номера страницы в фигуры на главной странице, чтобы текущий номер страницы отображается в верхней части каждой страницы.


```vb
Sub PageNumberShape() 
 With ActiveDocument.MasterPages(1).Shapes _ 
 .AddShape(Type:=msoShape5pointStar, Left:=36, _ 
 Top:=36, Width:=50, Height:=50) 
 With .TextFrame.TextRange 
 .InsertPageNumber 
 .ParagraphFormat.Alignment = pbParagraphAlignmentCenter 
 End With 
 .Fill.ForeColor.RGB = RGB(Red:=125, Green:=125, Blue:=255) 
 End With 
End Sub
```


