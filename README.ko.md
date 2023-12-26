# Raytracing in One Weekend, 한컴오피스 한셀 에디션

> 이 프로젝트는 [Project Lemonade](https://rangho.dev/project-lemonade)의
> 일부입니다.

한컴오피스 한셀의 VBScript 매크로로 작성된 간단한 레이트레이서입니다.
Peter Shirley 님의 책 [Raytracing in One Weekend][1]를 기반으로 작성했습니다.

[1]: https://raytracing.github.io/

## 한셀 파일 생성하기

먼저 [한컴오피스 한셀][2]이 설치되어 있는지 확인해 주세요.
한셀 파일 생성 스크립트와 매크로는 한컴오피스 한셀 2020 교육기관용 버전에서
테스트하였습니다.
최근 한컴오피스 웹 버전 출시 이후로 무료 체험판을 단종시켰습니다.
찾아 보면 아카이브된 설치 파일이 돌아다니기는 하지만, __이를 사용하는 것에 대하여
개발자는 어떠한 책임도 지지 않습니다__.

PowerShell을 열어서 이 레포지토리를 클론하시거나 소스코드를 다운받아 주세요.

```powershell
git clone https://github.com/RangHo/hancell-raytracing

# 아니면...

Invoke-WebRequest `
    "https://github.com/RangHo/hancell-raytracing/archive/refs/heads/main.zip" `
    -OutFile .\hancell-raytracing.zip
Expand-Archive .\hancell-raytracing.zip .\
Rename-Item .\hancell-raytracing-main .\hancell-raytracing
```

다운받은 소스코드 경로에 접근한 후, `Build.ps1` 스크립트를 실행해 주세요.

```powershell
Set-Location .\hancell-raytracing
.\Build.ps1
```

실행이 완료되면 `내 문서` 폴더에 `raytracing.cell` 파일이 생성됩니다.

만약 PowerShell이 스크립트를 검증할 수 없다는 오류를 뿜으면 아래 명령으로
스크립트를 실행할 수 있도록 설정해 주세요.

```powershell
Set-ExecutionPolicy `
    -Scope CurrentUser `
    -ExecutionPolicy Bypass
```

[2]: https://www.hancom.com/product/productWindowsMain.do?gnb0=23&gnb1=29

## 레이트레이싱 시작하기

생성된 스프레드시트 파일에는 아래 4개의 매크로가 저장되어 있습니다.

* `Init`: 렌더링용 캔버스 초기화
* `Render`: 레이트레이싱 시작
* `Benchmark`: 레이트레이싱 시작 및 성능 측정
* `Sketch`: VBScript 테스트용 매크로

먼저, 생성된 `raytracing.cell` 파일을 열어 주세요.

그 다음, `도구`, `매크로`, `매크로 실행` 메뉴를 차례로 선택한 후, 매크로 목록
창에서 `Init` 매크로를 실행합니다.
셀 크기와 기본 바탕색 설정이 완료되면, 같은 메뉴를 선택해서 `Render` 매크로를
실행해 주세요.

~~정신 나갈 것 같이~~ 오랜 시간이 지나고 나면 장면이 렌더링됩니다!

한셀 매크로는 설계상 싱글 스레드로 실행되기 때문에 매크로가 실행되는 도중에는
스프레드시트 윈도우와 어떤 상호작용도 할 수 없다는 점 주의하세요.
렌더링되는 과정을 확인하시려면 한셀 윈도우를 가장 위에 두고 다른 화면으로
전환하면 안 됩니다.
실수로 다른 화면을 켰다면 (한셀이 그 과정에서 최소화되지 않았다는 가정 하에)
한셀 위에 위치한 화면을 최소화해서 렌더링 과정을 다시 볼 수 있습니다.
