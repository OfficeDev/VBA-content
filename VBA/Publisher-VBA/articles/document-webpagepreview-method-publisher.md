---
title: "Метод Document.WebPagePreview (издатель)"
keywords: vbapb10.chm196724
f1_keywords: vbapb10.chm196724
ms.prod: publisher
api_name: Publisher.Document.WebPagePreview
ms.assetid: 44083fae-d21d-9cd3-3553-a4d4346141f5
ms.date: 06/08/2017
ms.openlocfilehash: 96f2b074015addb1c4b9ae3422980234be718415
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentwebpagepreview-method-publisher"></a>Метод Document.WebPagePreview (издатель)

Генерирует предварительный просмотр веб-страницы публикации, указанный в Internet Explorer.


## <a name="syntax"></a>Синтаксис

 _выражение_. **WebPagePreview**

 переменная _expression_A, представляющий объект **Document** .


## <a name="remarks"></a>Заметки

Веб-Просмотр могут создаваться для печати публикаций. Тем не менее внешний вид веб-Просмотр может отличаться от печати публикации.

Веб-Просмотр открывает active страницу. Предварительный просмотр веб-страницы создаются для каждой страницы публикации. Тем не менее если публикация является публикацией печати или в противном случае — не имеет панели навигации, может быть способ перехода к этих страниц.

Свойство **[PublicationType](document-publicationtype-property-publisher.md)** используется для определения, является ли публикации печати публикации или веб-публикации.

Этот метод соответствует команду **Предварительный просмотр веб-страницы** в меню **файл** .


## <a name="example"></a>Пример

В следующем примере задается активную страницу публикации и создает веб-Просмотр публикации.


```vb
 
With ActiveDocument 
 .ActiveView.ActivePage = .Pages(2) 
 .WebPagePreview 
End With
```


