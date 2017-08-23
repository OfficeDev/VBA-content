---
title: "Метод Pages.AddWizardPage (издатель)"
keywords: vbapb10.chm458758
f1_keywords: vbapb10.chm458758
ms.prod: publisher
api_name: Publisher.Pages.AddWizardPage
ms.assetid: c56db218-d0f4-4f13-dfde-6198dc63cc81
ms.date: 06/08/2017
ms.openlocfilehash: 207e1fe8f1e5fb2281dbb288b6766168228bb7c9
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pagesaddwizardpage-method-publisher"></a>Метод Pages.AddWizardPage (издатель)

Добавляет указанный новая страница мастера в указанное расположение в публикации.


## <a name="syntax"></a>Синтаксис

 _выражение_. **AddWizardPage** ( **_После_**, **_PageType_**, **_AddHyperlinkToWebNavBar_**)

 переменная _expression_A, представляет собой объект- **страниц** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|После|Обязательное свойство.| **Длинный**|Страница, после которого следует поместить новую страницу мастера.|
|PageType|Необязательный| **PbWizardPageType**|Тип страницы мастера, чтобы добавить.|
|AddHyperlinkToWebNavBar|Необязательный| **Boolean**|Указывает, добавляются ли ссылка на новую страницу на панели автоматического навигации существующих страниц. Значение по умолчанию — **False**, что означает, что если этот аргумент задан, ссылки на этой странице не добавляются на панели автоматического навигации существующих страниц.|

## <a name="remarks"></a>Заметки

Страницах мастера можно добавить только на аналогичную мастера публикации. К примеру можно добавить страница мастера календаря каталога в каталоге, но не на информационный бюллетень. При добавлении на страницу мастера в другой тип публикации, возникает ошибка.

Параметр PageType может иметь одно из **[PbWizardPageType](pbwizardpagetype-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.


## <a name="example"></a>Пример

В этом примере создается новая публикация каталога, добавляет страница мастера календаря после первой страницы каталога и добавляет страницы как ссылку к каждому набору панель навигации Web публикации.


```vb
Sub AddNewWizardPage() 
 Dim PubApp As Publisher.Application 
 Dim PubDoc As Publisher.Document 
 Set PubApp = New Publisher.Application 
 Set PubDoc = PubApp.NewDocument(Wizard:=pbWizardCatalogs, _ 
 Design:=7) 
 PubDoc.Pages.AddWizardPage After:=1, _ 
 PageType:=pbWizardPageTypeCatalogCalendar, _ 
 AddHyperLinkToWebNavBar:=True 
 PubApp.ActiveWindow.Visible = True 
End Sub
```

В этом примере выполняется проверка, что активный документ в каталоге и, если он установлен, добавляет каталога формы после первой страницы, но не добавляйте страницы как ссылку в наборах панели навигации веб.




```vb
Sub InsertCatalogWizardPage() 
 With ActiveDocument 
 If .Wizard.ID = 161 Then 
 .Pages.AddWizardPage After:=1, _ 
 PageType:=pbWizardPageTypeCatalogForm, _ 
 AddHyperLinkToWebNavBar:=False 
 End If 
 End With 
End Sub
```


