# gohwp
Go언어(Golang)을 이용한 한글 오토메이션 패키지

## 1. gohwp ?
* OLE Automation을 통해 한글(HWP) 자동화를 Go언어로 이용하기 위한 패키지입니다.
* Python을 이용한 HWP 자동화 예제는 많으나, Go언어를 이용한 예제가 없기에 시작하였습니다.
* 아래의 한글오토메이션의 API 매뉴얼에 나온 것들을 가급적 많이 구현하고자 하는 꿈이 있었으나, 개발이 업이 아닌 취미인지라 미구현된 부분이 많이 있습니다.
  * [ActionObject.pdf](https://github.com/hancom-io/devcenter-archive/raw/main/hwp-automation/ActionObject.pdf)
  * [HwpAutomation.pdf](https://github.com/hancom-io/devcenter-archive/raw/main/hwp-automation/HwpAutomation.pdf)
  * [ParameterSetObject.pdf](https://github.com/hancom-io/devcenter-archive/raw/main/hwp-automation/ParameterSetObject.pdf)

## 2. 시작하기

### 2.1. 설치

```bash
go get github.com/b612nightsky/gohwp
```

### 2.2. HWP 새창 띄우기
```go
package main

import (
	"github.com/b612nightsky/gohwp"
)

func main() {
  hwp := gohwp.Initialize()   // 초기화
  defer hwp.UnInitialize()
  hwp.ShowWindow(true)        // hwp문서창 보이기
  hwp.InsertText("안녕하세요") // 문자입력
}

```

## 3. 액션(Action) 사용

### 3.1. 파라미터 세트이 없어도 되는 액션

#### 3.1.1. Run 메서드
파라미터 값이 따로 있을 필요가 없는 액션들은 [ActionObject.pdf](https://github.com/hancom-io/devcenter-archive/raw/main/hwp-automation/ActionObject.pdf)에서 필요한 액션 이름을 찾아서 Run 메서드를 이용하면 됩니다.

예를 들어, 문서 처음으로 커서(캐럿)을 옮기는 액션은 다음과 같습니다.
```go
hwp.Run("MoveDocBegin")
```

#### 3.1.2. ActionObject
매번 액션 이름을 찾는 것도 귀찮고 액션 이름을 외우지 못하기에 [ActionObject.pdf](https://github.com/hancom-io/devcenter-archive/raw/main/hwp-automation/ActionObject.pdf)의 모든 액션은 메서드로 정의했습니다.

다만, 액션이 많기 때문에 기능별로 다음과 같이 구분지었습니다.
 * FrameActionWindow (창)
 * FrameActionEtc (기타)
 * FrameActionHelp (도움말)
 * FrameActionFile (파일)
 * FrameActionView (보기)
 * ActionFile (파일)
 * ActionEdit (편집)
 * ActionShape (모양)
 * ActionView (보기)
 * ActionInput (입력)
 * ActionTool (도구)
 * ActionTable (표)
 * ActionDraw (그리기)
 * ActionEtc (기타)

앞서 문서 처음으로 커서를 옮기는 예를 다음과 같이 이용할 수 있습니다.
```go
hwp.ActionObject().ActionEdit().MoveDocBegin()
```
이는 결국 Run 메서드를 이용하는 것이고, 다만 문자열로 된 액션 이름을 모두 외울 수 없기에 각각 메서드로 만들어 놓은 것입니다.

### 3.2. 파라미터 세트이 있어야 하는 액션
파라미터 세트이 있어야 하는 액션을 실행하는 방법은 아래와 같이 여러 방법이 있습니다. 아래 예제 모두 동일한 파라미터로 동일한 액션을 실행하고 있습니다.

#### 3.2.1. HAction


 * HWP 스크립트 매크로
 ```java
 HAction.GetDefault("Print", HParameterSet.HPrint.HSet); // 액션 초기화
 HParameterSet.HPrint.NumCopy = 3; //인쇄 매수를 3장 지정
 HAction.Execute("Print", HParameterSet.HPrint.HSet); // 액션 수행
 ```

 * gohwp
 ```go
 act := hwp.HAction() // 액션 생성
 para := hwp.HParameterSet().HPrint() // 파라미터 생성
 act.GetDefault("Print",para.HSet()) // 액션 & 파라미터 초기화
 para.NumCopy(3) // 인쇄 매수 3장 지정
 act.Execute("Print",para.HSet()) // 액션 수행
 ```

gohwp를 이용하면 작성해야하는 코드의 줄이 더 많아졌습니다만, 스크립트 매크로가 아닌 C++, C#에서 구현하려면 gohwp와 같이 HAction, HParameterSet 등의 변수를 선언해줘야 하기 때문에 크게 차이는 없습니다.

나중에 세부적인 파라미터를 직접 다루려면 결국 HAction을 이용하게 될 것입니다.

위 예시에서는 para.NumCopy(3)으로 값을 입력하였는데, 인자를 전달하지 않고 para.NumCopy()를 하면 현재 NumCopy에 해당하는 값을 반환합니다.

#### 3.2.2. IDHwpAction 

한컴 오토메이션에는 사용자에게 편의를 제공하는 클래스 및 함수가 있습니다. IDHwpAction도 그 중 하나로 이를 이용하면 앞서 HAction과 동일한 것을 다음과 같이 이용할 수 있습니다.

```go
act := hwp.CreateAction("Print")
para := act.CreateSet()
act.GetDefault(para)
para.SetItem("NumCopy",3)
act.Execute(para)|
```

HAction을 사용하는 것과 비교하면 어떤 파라미터 세트을 생성해야하는지 몰라도 되는데, 단점으로는 파라미터의 이름을 문자열로 알고 있어야 합니다. 

#### 3.2.3. ActionWithParameters() 메서드 이용

IDHwpAction을 이용하더라도 코드를 여러 줄을 작성해야 합니다. 이마저도 귀찮아서 한 줄로 작성하게끔 메서드를 만든 것이 ActionWithParameters입니다.

```go
hwp.ActionWithParameters("Print","NumCopy",3)
```
위의 IDHwpAction을 이용하는 방법을 한줄로 줄여놓은 메서드입니다. 액션 이름과 파라미터 이름을 문자열로 알고 있어야 합니다. 

단점으로는 간혹 파라미터가 배열(HArrar)인 경우에는 이 방법으로 입력할 수 없습니다. HArray의 경우 배열을 생성하고 배열 내 각각 위치에 값을 입력하는 행위를 해야하는데, 이를 하려면 HAction에서 하는 것이 낫습니다.


```go
hwp.ActionWithParameters("액션이름","파라미터이름1",값1,"파라미터이름2","값2",...) 
```
방법으로 한번에 여러 파라미터를 같이 인자로 넘겨서 실행시킵니다.

#### 3.2.4. ActionObject
앞서 파라미터가 없는 액션을 다루는 방법과 바로 위 ActionWithParameters를 합쳐놓은 방법입니다.
```go
hwp.ActionObject().ActionFile().Print("Copy",3)
```
이 방법 역시 파라미터 이름은 문자열로 알고 있어야 합니다.

## 주의
※ gohwp에서 발생하는 error는 전부 panic으로 처리하고 있습니다.

일반적으로 Go언어에서는 panic을 지양한다고 합니다만, gohwp는 문제가 생겼을 때 에러를 던져주지 않고 panic을 발생시키고 있습니다. 이에 대해서는 많은 고민이 있었습니다. 마치 스크립트 언어처럼 HWP 자동화 코드를 작성하고 싶은데 매번 에러가 nil이 아님을 확인하기는 번거로웠습니다. 물론 사용자 측에서 에러 값을 받지 않고 확인하지 않는 방법도 있긴 합니다. 하지만, 원하는 문서 편집에 문제가 생기면 panic으로 인지하고 중단하는 것 택했습니다. 계속 고민하고 있는 부분으로 에러 처리 방법을 변경하게 된다면 그땐 패키지가 크게 변하지 않을까 싶습니다.