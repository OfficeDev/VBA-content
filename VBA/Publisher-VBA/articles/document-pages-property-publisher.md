---
title: "Свойство Document.Pages (издатель)"
keywords: vbapb10.chm196631
f1_keywords: vbapb10.chm196631
ms.prod: publisher
api_name: Publisher.Document.Pages
ms.assetid: 2bb3e529-a459-b37c-c9ae-4cc059954a63
ms.date: 06/08/2017
ms.openlocfilehash: f8f15ed8afa1ba06ea89e47907de84ee4b9f9543
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentpages-property-publisher"></a>Свойство Document.Pages (издатель)

Возвращает набор **[страниц](pages-object-publisher.md)** , представляющий все страницы в указанной публикации.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Страницы**

 переменная _expression_A, представляющий объект **Document** .


## <a name="example"></a>Пример

Следующий пример возвращает набор **страниц** публикации, активных и отчетов, сколько страниц.


```vb
Dim pgsTemp As Pages 
 
Set pgsTemp = ActiveDocument.Pages 
 
With pgsTemp 
 MsgBox "There are " &; .Count _ 
 &; " page(s) in the active publication." 
End With
```


