---
title: "Свойство Section.ShowHeaderFooterOnFirstPage (издатель)"
keywords: vbapb10.chm7405574
f1_keywords: vbapb10.chm7405574
ms.prod: publisher
api_name: Publisher.Section.ShowHeaderFooterOnFirstPage
ms.assetid: 6c814884-9bee-72ae-3a40-5118bebd6f02
ms.date: 06/08/2017
ms.openlocfilehash: bbbb4977545e5aafee6a7df784372246dd198edc
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="sectionshowheaderfooteronfirstpage-property-publisher"></a>Свойство Section.ShowHeaderFooterOnFirstPage (издатель)

 **Значение true,** Если заголовок и нижний колонтитул из указанного раздела будут отображаться. Чтение и запись **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ShowHeaderFooterOnFirstPage**

 переменная _expression_A, представляет собой объект **раздела** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="example"></a>Пример

В следующем примере добавляется новый раздел, начиная на второй странице активного документа добавляется текст колонтитулов на главную страницу и затем задает для свойства **ShowHeaderFooterOnFirstPage** значение **True**.


```vb
Dim objSection As Section 
Set objSection = ActiveDocument.Sections.Add(StartPageIndex:=2) 
With ActiveDocument.Pages(2).Master 
 .Header.TextRange.Text = "Page " &; .PageNumber &; " header." 
 .Footer.TextRange.Text = "Page " &; .PageNumber &; " footer." 
End With 
objSection.ShowHeaderFooterOnFirstPage = True
```


