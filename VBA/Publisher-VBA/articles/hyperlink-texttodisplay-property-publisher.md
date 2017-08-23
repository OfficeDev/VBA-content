---
title: "Свойство Hyperlink.TextToDisplay (издатель)"
keywords: vbapb10.chm4587536
f1_keywords: vbapb10.chm4587536
ms.prod: publisher
api_name: Publisher.Hyperlink.TextToDisplay
ms.assetid: 26b5857c-3f94-0d33-f65e-9c34f2a4cc2b
ms.date: 06/08/2017
ms.openlocfilehash: e2f0b2bc5c28c75adc927c5aeca19aa590895a56
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="hyperlinktexttodisplay-property-publisher"></a>Свойство Hyperlink.TextToDisplay (издатель)

Возвращает или задает **строку** , которая представляет собой текст, отображаемый для гиперссылки. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **TextToDisplay**

 переменная _expression_A, представляющий объект **гиперссылки** .


### <a name="return-value"></a>Возвращаемое значение

String


## <a name="example"></a>Пример

В этом примере задается текст гиперссылки и адрес гиперссылки на первой странице. В этом примере предполагается, что первая страница активная публикация содержит по крайней мере один фигуры с по крайней мере один текст гиперссылки.


```vb
Sub SetHyperlinkTextToDisplay() 
 With ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.Hyperlinks.Item(1) 
 .TextToDisplay = "Tailspin Toys Web Site" 
 .Address = "http://www.tailspintoys.com/" 
 End With 
End Sub
```


