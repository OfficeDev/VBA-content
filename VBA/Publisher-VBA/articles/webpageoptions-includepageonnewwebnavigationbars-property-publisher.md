---
title: "Свойство WebPageOptions.IncludePageOnNewWebNavigationBars (издатель)"
keywords: vbapb10.chm544773
f1_keywords: vbapb10.chm544773
ms.prod: publisher
api_name: Publisher.WebPageOptions.IncludePageOnNewWebNavigationBars
ms.assetid: 5e2f60d0-e812-8ca1-e54b-33a1f9eedf84
ms.date: 06/08/2017
ms.openlocfilehash: 101e74abfb316856a040dbe43521e8eff2a44b6d
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webpageoptionsincludepageonnewwebnavigationbars-property-publisher"></a>Свойство WebPageOptions.IncludePageOnNewWebNavigationBars (издатель)

Возвращает или задает **логическое** значение, указывающее, добавляются ли ссылка на веб-страницы на панели автоматического навигации новых страниц. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **IncludePageOnNewWebNavigationBars**

 переменная _expression_A, представляющий объект **WebPageOptions** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

Значение по умолчанию свойства **IncludePageOnNewWebNavigationBars** — **значение False**, это означает, что ссылки на указанной странице не добавляются на панели автоматического навигации новых страниц.

Назначить этому свойству значение **False,** не удаляйте ссылки на страницу указанного из любого панелей автоматического навигации, которые уже их, но его запретить ссылки на страницы добавляются на панели навигации автоматического новых страниц.

Установка для этого свойства **значения True** применяется только к панелей навигации автоматического новых страниц и не обновляет существующий панелей автоматического навигации в веб-публикации.

При добавлении новой страницы в веб-публикации с помощью ** [Pages.Add](pages-add-method-publisher.md)** метод, необязательный параметр **AddHyperlinkToWebNavBar** можно использовать для указания ли существующих панелей навигации автоматического будут добавлены ссылки на странице "новый". Значение этого параметра используется для заполнения значение свойства **IncludePageOnNewWebNavigationBars** .


## <a name="example"></a>Пример

Следующий пример указывает, что позволяет использовать второй active веб-публикации страницы будет добавлен на панели автоматического навигации новых страниц. Обратите внимание, что если добавляется новая страница публикации после этого момента, свойство **IncludePageOnNewWebNavigationBars** будет иметь **значение False**.


```vb
Dim theWPO As WebPageOptions 
 
Set theWPO = ActiveDocument.Pages(2).WebPageOptions 
With theWPO 
 .IncludePageOnNewWebNavigationBars = True 
End With
```

В следующем примере показано добавление двух новых страниц в публикации с помощью метода **Pages.Add** . Параметр **AddHyperlinkToWebNavBar** имеет значение **True**, указывает, позволяет использовать эти две новые страницы добавляются панелей навигации автоматического существующих страниц.

Затем добавляется другой страницы публикации, а **AddHyperlinkToWebNavBar** задан. Это означает, что свойство **IncludePageOnNewWebNavigationBars** имеет **значение False** для только что добавленный страницы и ссылки на этой странице не будут включены в панели автоматического переходов существующих страниц.




```vb
Dim thePage As page 
Dim thePage2 As page 
 
Set thePage = ActiveDocument.Pages.Add(Count:=2, _ 
 After:=4, AddHyperlinkToWebNavBar:=True) 
 
Set thePage2 = ActiveDocument.Pages.Add(Count:=1, After:=6)
```


