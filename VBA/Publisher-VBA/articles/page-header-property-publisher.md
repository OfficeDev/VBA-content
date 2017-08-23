---
title: "Свойство Page.Header не имело значения (издатель)"
keywords: vbapb10.chm393247
f1_keywords: vbapb10.chm393247
ms.prod: publisher
api_name: Publisher.Page.Header
ms.assetid: f10806eb-972a-d482-935c-95d5ccbbbb36
ms.date: 06/08/2017
ms.openlocfilehash: 048b88d7a192ac8fc4607037b0322a902ef6599d
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pageheader-property-publisher"></a>Свойство Page.Header не имело значения (издатель)

Возвращает объект **HeaderFooter** , представляющий заголовок на указанный объект **страницы** . Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Заголовок**

 переменная _expression_A, представляющий объект **Page** .


### <a name="return-value"></a>Возвращаемое значение

HeaderFooter


## <a name="remarks"></a>Заметки

Это свойство используется только для главных страниц. Будет возвращена ошибка «Эта возможность предназначена только для главных страниц», если свойство заголовка можно обратиться из объекта **Page** , который возвращается форме коллекции **Pages** . Новый объект **HeaderFooter** создается для указанного главной страницы, доступ к этому свойству.


## <a name="example"></a>Пример

В следующем примере создается объект **HeaderFooter** и задает верхний колонтитул первой главной страницы.


```vb
Dim objHeader As HeaderFooter 
Set objHeader = ActiveDocument.MasterPages(1).Header
```

**HeaderFooter** объект, возвращенный свойством **заголовка** можно использовать для работы с содержимое заголовка. В следующем примере задается некоторые свойства объекта **HeaderFooter** первой главной страницы.




```vb
With ActiveDocument.masterPages(1) 
 With .Header 
 .TextRange.Text = "Windows" &; Chr(13) &; "Office" &; Chr(13) &; "Internet Explorer" 
 With .TextRange.ParagraphFormat 
 .SetListType Value:=pbListTypeBullet, BulletText:="*" 
 .Alignment = pbParagraphAlignmentLeft 
 End With 
 End With 
 With .Footer 
 .TextRange.Hyperlinks.Add Text:=.TextRange, _ 
 Address:="http://www.tailspintoys.com", _ 
 TextToDisplay:="Tailspin" 
 End With 
End With
```


