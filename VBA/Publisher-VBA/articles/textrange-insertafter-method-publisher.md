---
title: "Метод TextRange.InsertAfter (издатель)"
keywords: vbapb10.chm5308448
f1_keywords: vbapb10.chm5308448
ms.prod: publisher
api_name: Publisher.TextRange.InsertAfter
ms.assetid: f647be29-68c7-b221-adf1-fa233583e74e
ms.date: 06/08/2017
ms.openlocfilehash: afe30fcfbd578199c97a1937dbfc03d416319ba2
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textrangeinsertafter-method-publisher"></a>Метод TextRange.InsertAfter (издатель)

Возвращает объект **[TextRange](textrange-object-publisher.md)** , представляющий текст, добавляемый в конец диапазона текста.


## <a name="syntax"></a>Синтаксис

 _выражение_. **InsertAfter** ( **_NewText_**)

 переменная _expression_A, представляющий объект **TextRange** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|NewText|Обязательное свойство.| **String**|Текст для вставки.|

### <a name="return-value"></a>Возвращаемое значение

TextRange


## <a name="example"></a>Пример

В этом примере добавляется номер сборки Microsoft Publisher в конец первую фигуру на первой странице active публикации. В этом примере предполагается, что указанные форму — фрагмент текста и не другого типа фигуры.


```vb
Sub AppendText() 
 With ActiveDocument.Pages(1).Shapes(1) 
 .TextFrame.TextRange.InsertAfter _ 
 NewText:="Microsoft Publisher Build : " &; Build 
 End With 
End Sub
```


