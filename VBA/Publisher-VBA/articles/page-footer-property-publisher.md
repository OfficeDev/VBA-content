---
title: "Свойство Page.Footer (издатель)"
keywords: vbapb10.chm393248
f1_keywords: vbapb10.chm393248
ms.prod: publisher
api_name: Publisher.Page.Footer
ms.assetid: 8ab5a59b-c8d5-6217-098c-c53336ee5311
ms.date: 06/08/2017
ms.openlocfilehash: 4182a0626790cd5f4f3130fef971e5784493bd4f
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pagefooter-property-publisher"></a>Свойство Page.Footer (издатель)

Возвращает объект **HeaderFooter** , представляющий нижнего колонтитула на указанный объект **страницы** . Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Нижний колонтитул**

 переменная _expression_A, представляющий объект **Page** .


### <a name="return-value"></a>Возвращаемое значение

HeaderFooter


## <a name="remarks"></a>Заметки

Это свойство используется только для главных страниц. «Эта возможность предназначена только для главной страницы» ошибка возвращается в том случае, если нижний колонтитул, свойство можно обратиться из объекта **Page** , возвращенный формы коллекции **Pages** . Новый объект **HeaderFooter** создается для указанного главной страницы, доступ к этому свойству.


## <a name="example"></a>Пример

В следующем примере создается объект **HeaderFooter** и задает нижний колонтитул первой главной страницы.


```vb
Dim objFooter As HeaderFooter 
Set objFooter = ActiveDocument.MasterPages(1).Footer
```

**HeaderFooter** объект, возвращенный свойством **нижнего колонтитула** можно использовать для управления содержимым нижнего колонтитула. В следующем примере задается некоторые свойства объекта **HeaderFooter** первой главной страницы.




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


