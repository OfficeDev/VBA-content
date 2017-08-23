---
title: "Метод FindReplace.Execute (издатель)"
keywords: vbapb10.chm8323086
f1_keywords: vbapb10.chm8323086
ms.prod: publisher
api_name: Publisher.FindReplace.Execute
ms.assetid: 351a64ab-3c6c-c9c9-7ffe-b60b73d390ae
ms.date: 06/08/2017
ms.openlocfilehash: be53f1680913d5f84ba8ded0fd8b40814ef42a72
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="findreplaceexecute-method-publisher"></a>Метод FindReplace.Execute (издатель)

Выполняет указанной операции поиска и замены.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Выполнение**

 переменная _expression_A, представляет собой объект- **FindReplace** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="example"></a>Пример

В этом примере выполняет операцию поиска и замены активных документов.


```vb
Sub ExecuteFindReplace() 
 Dim objFindReplace As FindReplace 
 Set objFindReplace = ActiveDocument.Find 
 With objFindReplace 
 .Clear 
 .FindText = "library" 
 .Execute 
 End With 
End Sub
```


