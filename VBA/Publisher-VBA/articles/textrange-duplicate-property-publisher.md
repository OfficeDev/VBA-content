---
title: "Свойство TextRange.Duplicate (издатель)"
keywords: vbapb10.chm5308466
f1_keywords: vbapb10.chm5308466
ms.prod: publisher
api_name: Publisher.TextRange.Duplicate
ms.assetid: 545dbfdb-4cd5-99b1-1ba3-b723e8d7b827
ms.date: 06/08/2017
ms.openlocfilehash: 3512f75059851f0ba8874ea4169c343ad20035f2
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textrangeduplicate-property-publisher"></a>Свойство TextRange.Duplicate (издатель)

Возвращает объект **[TextRange](textrange-object-publisher.md)** , представляющий дублировать диапазон указанный текст.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Дублирующиеся**

 переменная _expression_A, представляющий объект **TextRange** .


### <a name="return-value"></a>Возвращаемое значение

TextRange


## <a name="example"></a>Пример

В этом примере задается значение строковой переменной содержимого указанного текстового поля на первой странице active публикации. Затем создается новая страница с текстовое поле и задает содержимое новое текстовое поле равно значению строковой переменной.


```vb
Sub DuplicateTextBoxContents() 
 Dim strDuplicate As String 
 Dim pagNew As Page 
 
 With ThisDocument.Pages(1).Shapes(1).TextFrame.TextRange 
 strDuplicate = .Duplicate 
 End With 
 
 Set pagNew = ThisDocument.Pages.Add(Count:=1, After:=1) 
 
 pagNew.Shapes.AddTextbox(Orientation:=pbTextOrientationHorizontal, _ 
 Left:=72, Top:=72, Width:=200, Height:=200).TextFrame _ 
 .TextRange.Text = strDuplicate 
End Sub
```


