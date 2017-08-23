---
title: "Свойство Page.IsTwoPageMaster (издатель)"
keywords: vbapb10.chm131100
f1_keywords: vbapb10.chm131100
ms.prod: publisher
api_name: Publisher.Page.IsTwoPageMaster
ms.assetid: dbfc3c21-0070-3f0a-c0b0-746d83c46765
ms.date: 06/08/2017
ms.openlocfilehash: 566044fadb59a74375058a48d2f2aee963774cda
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pageistwopagemaster-property-publisher"></a>Свойство Page.IsTwoPageMaster (издатель)

 **Значение true,** Если указанный объект **Page** является основным две страницы. Чтение и запись **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **IsTwoPageMaster**

 переменная _expression_A, представляющий объект **страницы** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

Этот метод работает только главных страниц. Возвращает ошибку **Эта возможность предназначена только для главных страниц** при попытке доступа к этому свойству из объекта страницы публикации.


## <a name="example"></a>Пример

В следующем примере добавляется текст для каждого заголовка главной страницы две страницы, указание главной страницы PageNumber и его место в распространении: 1 или 2.


```vb
Dim objMasterPage As Page 
Dim pageCount As Long 
Dim i As Long 
pageCount = ActiveDocument.MasterPages.Count 
For i = 1 To pageCount 
 Set objMasterPage = ActiveDocument.MasterPages(i) 
 If objMasterPage.IsTwoPageMaster Then 
 objMasterPage.Header.TextRange.Text = "MasterPage " &; _ 
 objMasterPage.PageNumber &; ", Page 1 of 2" 
 i = i + 1 
 Set objMasterPage = ActiveDocument.MasterPages(i) 
 objMasterPage.Header.TextRange.Text = "MasterPage " &; _ 
 objMasterPage.PageNumber &; ", Page 2 of 2" 
 End If 
Next i 

```


