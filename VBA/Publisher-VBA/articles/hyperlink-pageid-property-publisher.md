---
title: "Свойство Hyperlink.PageID (издатель)"
keywords: vbapb10.chm4587525
f1_keywords: vbapb10.chm4587525
ms.prod: publisher
api_name: Publisher.Hyperlink.PageID
ms.assetid: 1b5051eb-e6b4-a5a7-610a-5be03863a92b
ms.date: 06/08/2017
ms.openlocfilehash: 559fb42cd9e607f57ca0dd4b8af574305862315d
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="hyperlinkpageid-property-publisher"></a>Свойство Hyperlink.PageID (издатель)

Возвращает или задает **Long** , указывающее, страницы публикации, который является целевым для указанного гиперссылки. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PageID**

 переменная _expression_A, представляющий объект **гиперссылки** .


## <a name="example"></a>Пример

В следующем примере сообщается какие гиперссылки в активной публикации, связанной со страницы.


```vb
Dim hypTemp As Hyperlink 
Dim lngID As Long 
Dim strPage As String 
 
Set hypTemp = ActiveDocument.Pages(1).Shapes(1).Hyperlink 
 
lngID = hypTemp.PageID 
strPage = ActiveDocument.Pages.FindByPageID(PageID:=lngID).PageNumber 
 
MsgBox "This hyperlink goes to the page " &; strPage &; "."
```


