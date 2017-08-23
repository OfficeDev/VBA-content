---
title: "Метод Document.Close (издатель)"
keywords: vbapb10.chm196680
f1_keywords: vbapb10.chm196680
ms.prod: publisher
api_name: Publisher.Document.Close
ms.assetid: b4b21484-1858-b7b3-291f-18ef8cab8ba7
ms.date: 06/08/2017
ms.openlocfilehash: 16f0153169d7ab62e39435933657b828120ac348
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentclose-method-publisher"></a>Метод Document.Close (издатель)

Закрывает текущий публикации и создается пустой публикации вместо него.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Закрытие**

 переменная _expression_A, представляющий объект **Document** .


## <a name="remarks"></a>Заметки

Метод **Close** можно использовать только на объект открытых **документов** в другой экземпляр Microsoft Publisher. Закрывается active публикации в текущем экземпляре Publisher приводит к ошибке.


## <a name="example"></a>Пример

В этом примере открывает публикации в новый экземпляр объекта Publisher для внесения изменений и затем закрывает публикации. (Обратите внимание, что этот пример работал, необходимо заменить _имя файла_ на допустимое имя файла).


```vb
Sub ModifyAnotherPublication() 
 
 ' Create new instance of Publisher. 
 Dim appPub As New Publisher.Application 
 
 ' Open publication. 
 appPub.Open FileName:="Filename" 
 
 ' Put code here to modify the publication as necessary. 
 
 ' Close the publication. 
 appPub.ActiveDocument.Close 
 
 ' Release the other instance of Publisher. 
 Set appPub = Nothing 
 
End Sub
```


