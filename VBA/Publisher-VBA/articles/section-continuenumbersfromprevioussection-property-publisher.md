---
title: "Свойство Section.ContinueNumbersFromPreviousSection (издатель)"
keywords: vbapb10.chm7405575
f1_keywords: vbapb10.chm7405575
ms.prod: publisher
api_name: Publisher.Section.ContinueNumbersFromPreviousSection
ms.assetid: a3d64f14-dc65-4fb1-5079-0fdf2e3f8f38
ms.date: 06/08/2017
ms.openlocfilehash: 97ceee6e427ec51d15e895907e6a919bb1f3541b
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="sectioncontinuenumbersfromprevioussection-property-publisher"></a>Свойство Section.ContinueNumbersFromPreviousSection (издатель)

 **Значение true,** Если указанный раздел по-прежнему производится нумерации из раздела prvious. Чтение и запись **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ContinueNumbersFromPreviousSection**

 переменная _expression_A, представляет собой объект **раздела** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="example"></a>Пример

В следующем примере добавляет три страницы публикации, добавляет новый раздел после первой страницы и затем задает **ContinueNumbersFromPreviousSection** значение **False** для нового раздела.


```vb
Dim objSection As Section 
ActiveDocument.Pages.Add Count:=3, After:=1 
Set objSection = ActiveDocument.Sections.Add(StartPageIndex:=2) 
objSection.ContinueNumbersFromPreviousSection = False 
 
 

```


