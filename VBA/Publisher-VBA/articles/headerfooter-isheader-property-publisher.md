---
title: "Свойство HeaderFooter.IsHeader (издатель)"
keywords: vbapb10.chm7471109
f1_keywords: vbapb10.chm7471109
ms.prod: publisher
api_name: Publisher.HeaderFooter.IsHeader
ms.assetid: b652fcc8-2c89-6d4f-c366-4c78681bea59
ms.date: 06/08/2017
ms.openlocfilehash: acf1c252d2591b35d189df51a4400085e9ad99a0
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="headerfooterisheader-property-publisher"></a>Свойство HeaderFooter.IsHeader (издатель)

 **Значение true,** Если указанный объект **HeaderFooter** заголовок, **значение False** , если это нижнего колонтитула. Только для чтения **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **IsHeader**

 переменная _expression_A, представляющий объект **HeaderFooter** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="example"></a>Пример

В следующем примере создается коллекция и заполняет его верхних и нижних колонтитулов из каждого главной страницы. Затем итерация коллекции и выполняется тест для определения ли объект **HeaderFooter** является верхних и нижних колонтитулов, а затем соответствующий текст записывается верхнего или нижнего колонтитула.


```vb
Dim objHeadersFooters As Collection 
Dim objMasterPage As page 
Dim objHeadFoot As HeaderFooter 
 
Set objHeadersFooters = New Collection 
For Each objMasterPage In ActiveDocument.masterPages 
 objHeadersFooters.Add objMasterPage.Header 
 objHeadersFooters.Add objMasterPage.Footer 
Next objMasterPage 
 
For Each objHeadFoot In objHeadersFooters 
 If objHeadFoot.IsHeader = True Then 
 objHeadFoot.TextRange.Text = "Header Text" 
 ElseIf objHeadFoot.IsHeader = False Then 
 objHeadFoot.TextRange.Text = "Footer Text" 
 End If 
Next 

```


