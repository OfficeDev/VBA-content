---
title: "Метод TextRange.InsertMailMergeField (издатель)"
keywords: vbapb10.chm5308483
f1_keywords: vbapb10.chm5308483
ms.prod: publisher
api_name: Publisher.TextRange.InsertMailMergeField
ms.assetid: 97bce07d-b831-3ad6-2436-f85590c3bcd8
ms.date: 06/08/2017
ms.openlocfilehash: 2820d136c0b38bed551d4cd425ec6fc91a589549
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textrangeinsertmailmergefield-method-publisher"></a>Метод TextRange.InsertMailMergeField (издатель)

Возвращает объект **[TextRange](textrange-object-publisher.md)** , представляющий текстового поля данных для слияния почты и объединение в каталог.


## <a name="syntax"></a>Синтаксис

 _выражение_. **InsertMailMergeField** ( **_varIndex_**)

 переменная _expression_A, представляющий объект **TextRange** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|varIndex|Обязательное свойство.| **Variant**|Имя или индекс поля данных в источнике данных.|

### <a name="return-value"></a>Возвращаемое значение

TextRange


## <a name="remarks"></a>Заметки

Для публикации области объединения в каталог для хранения данных текстовых полей он должен содержать по крайней мере один текстовое поле для хранения данных текстовых полей. 


## <a name="example"></a>Пример

В этом примере Вставка поля **LastName** в позиции курсора. В этом примере предполагается, что активная публикация является публикацией слияния почты и где-нибудь является позиции курсора в текстовом поле.


```vb
Sub InsertMergeField() 
 Selection.TextRange.InsertMailMergeField varIndex:="LastName" 
End Sub
```

В этом примере добавляется текстовое поле область указанной публикации и вставляет текстового поля данных в текстовом поле. В этом примере предполагается, что указанной публикации подключен к источнику данных, и что он содержит области объединения в каталог.




```vb
Set pbTextBox1 = ThisDocument.Pages(1).Shapes.AddTextbox(1, 100, 100, 175, 25) 
pbTextBox1.AddToCatalogMergeArea 
 
With pbTextBox1.TextFrame.TextRange 
 .Text = "List Price: " 
 .InsertMailMergeField "List Price" 
End With 

```


