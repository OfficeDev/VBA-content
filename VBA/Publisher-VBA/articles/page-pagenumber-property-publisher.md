---
title: "Свойство Page.PageNumber (издатель)"
keywords: vbapb10.chm393220
f1_keywords: vbapb10.chm393220
ms.prod: publisher
api_name: Publisher.Page.PageNumber
ms.assetid: 670e3f46-9cad-b85e-b627-3be8c7c4e577
ms.date: 06/08/2017
ms.openlocfilehash: 5a9270a3b3482c58b706f0113dfb5f09cfc8f860
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pagepagenumber-property-publisher"></a>Свойство Page.PageNumber (издатель)

Возвращает **строку** , представляющую текущий номер страницы. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PageNumber**

 переменная _expression_A, представляющий объект **Page** .


### <a name="return-value"></a>Возвращаемое значение

String


## <a name="example"></a>Пример

В этом примере создается текстовое поле, возвращает номер текущей страницы и вставляет его с помощью нового текста в его.


```vb
Sub GetPageNumber() 
 Dim strPageNumber As String 
 With ActiveDocument.Pages(1) 
 strPageNumber = .PageNumber 
 .Shapes.AddTextbox(Orientation:=pbTextOrientationHorizontal, _ 
 Left:=100, Top:=100, Width:=100, Height:=100) _ 
 .TextFrame.TextRange.InsertAfter NewText:="This is page " _ 
 &; strPageNumber &; " of " &; .Parent.Count &; "." 
 End With 
End Sub
```


