---
title: "Свойство Document.PageSetup (издатель)"
keywords: vbapb10.chm196632
f1_keywords: vbapb10.chm196632
ms.prod: publisher
api_name: Publisher.Document.PageSetup
ms.assetid: 1dac39f0-2507-a85b-8c71-cd1980022fb3
ms.date: 06/08/2017
ms.openlocfilehash: 2ea9e0543ec67f9c23ef60f692ea1b74e06c69ad
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentpagesetup-property-publisher"></a>Свойство Document.PageSetup (издатель)

**[PageSetup](pagesetup-object-publisher.md)** возвращает объект, представляющий размер страницы публикации, макет страницы и параметры бумаги. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PageSetup**

 переменная _expression_A, представляющий объект **Document** .


### <a name="return-value"></a>Возвращаемое значение

PageSetup


## <a name="remarks"></a>Заметки

Свойство **PageSetup** можно использовать только при печати нескольких страниц на одном листе бумаги. Если размер страницы больше половины размер бумаги, будут отображены ошибки.


## <a name="example"></a>Пример

В этом примере задает параметры страницы для публикации на нескольких страницах публикации на каждом листе бумаги при выводе на печать.


```vb
Sub SetTopMargin() 
 With ActiveDocument.PageSetup 
 .PageHeight = InchesToPoints(5) 
 .PageWidth = InchesToPoints(8) 
 .MultiplePagesPerSheet = True 
 .TopMargin = InchesToPoints(0.25) 
 .LeftMargin = InchesToPoints(0.25) 
 End With 
End Sub
```


