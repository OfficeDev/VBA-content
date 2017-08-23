---
title: "Метод TextRange.InsertBefore (издатель)"
keywords: vbapb10.chm5308449
f1_keywords: vbapb10.chm5308449
ms.prod: publisher
api_name: Publisher.TextRange.InsertBefore
ms.assetid: b0e4355b-b1bc-ae78-08ad-000d577fd7db
ms.date: 06/08/2017
ms.openlocfilehash: d19f7c24c456214e16a27fe8e61aa77924209f8f
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textrangeinsertbefore-method-publisher"></a>Метод TextRange.InsertBefore (издатель)

Добавляет новую строку в начало диапазона указанный текст. Возвращает объект **TextRange** , представляющий добавленный текст. При использовании без аргумента, этот метод возвращает пустую строку в конце указанного диапазона.


## <a name="syntax"></a>Синтаксис

 _выражение_. **InsertBefore** ( **_NewText_**)

 переменная _expression_A, представляющий объект **TextRange** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|NewText|Обязательное свойство.| **String**|Текст для вставки. Значение по умолчанию — пустая строка.|

### <a name="return-value"></a>Возвращаемое значение

TextRange


## <a name="example"></a>Пример

В этом примере добавляет номер сборки Microsoft Publisher и конца абзаца в начало первой фигуры на первой странице active публикации. В этом примере предполагается, что указанные форму — фрагмент текста и не другого типа фигуры.


```vb
Sub InsertTextBefore() 
 With ActiveDocument.Pages(1).Shapes(1) 
 .TextFrame.TextRange.InsertBefore _ 
 NewText:="Microsoft Publisher Build : " &; Build &; vbCrLf 
 End With 
End Sub
```


