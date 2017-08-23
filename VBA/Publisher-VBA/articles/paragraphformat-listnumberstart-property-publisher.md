---
title: "Свойство ParagraphFormat.ListNumberStart (издатель)"
keywords: vbapb10.chm5439527
f1_keywords: vbapb10.chm5439527
ms.prod: publisher
api_name: Publisher.ParagraphFormat.ListNumberStart
ms.assetid: 8e17fdaa-f53e-26c4-d92b-8ead65c28555
ms.date: 06/08/2017
ms.openlocfilehash: c9ce386df7b7456647c5b00b77f9ae4d1fd0080a
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="paragraphformatlistnumberstart-property-publisher"></a>Свойство ParagraphFormat.ListNumberStart (издатель)

Задает или получает **времени** , представляющий начальный номер списка. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ListNumberStart**

 переменная _expression_A, представляет собой объект- **ParagraphFormat** .


### <a name="return-value"></a>Возвращаемое значение

Длинный


## <a name="remarks"></a>Заметки

Возвращает сообщение «Доступ запрещен», если список не нумерованного списка.


## <a name="example"></a>Пример

В этом примере задается тип списка объекта **ParagraphFormat** **pbListTypeArabic** и устанавливает для свойства **ListNumber** значение 4.


```vb
Dim objParaForm As ParagraphFormat 
 
Set objParaForm = ActiveDocument.Pages(1).Shapes(1) _ 
.TextFrame.TextRange.ParagraphFormat 
 
 With objParaForm 
 .SetListType pbListTypeArabic 
 .ListNumberStart = 4 
 End With 
 
End Sub
```


