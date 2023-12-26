# Raytracing in One Weekend, Hancom Office Hancell edition

<p align="center">
    &middot;
    <a href="README.ko.md">한국어</a>
    &middot;
</p>

> This project is part of [Project Lemonade](https://rangho.dev/project-lemonade).

Simple path tracer written in Hancom Office Hancell VBScript, based on
[Raytracing in One Weekend][1] online book.

[1]: https://raytracing.github.io/

## Build instruction

Make sure that [Hancom Office Hancell][2] is installed.
The build script and macro code themselves are tested on Hancom Office Hancell
2020 for Educational Institutions.
Hancom recently shut down their free trial of Hancom Office in favor of the web
version, but there always exists a third-party archive.
__USE THAT AT YOUR OWN RISK.__

Start PowerShell and clone or download the repository source code.

```powershell
git clone https://github.com/RangHo/hancell-raytracing

# or...

Invoke-WebRequest `
    "https://github.com/RangHo/hancell-raytracing/archive/refs/heads/main.zip" `
    -OutFile .\hancell-raytracing.zip
Expand-Archive .\hancell-raytracing.zip .\
Rename-Item .\hancell-raytracing-main .\hancell-raytracing
```

Enter the source code directory and run the `Build.ps1` script.

```powershell
Set-Location .\hancell-raytracing
.\Build.ps1
```

Now a file named `raytracing.cell` should be available in your `Documents` folder.

If PowerShell whines that it cannot verify the script's authenticity, then make
it shut up with the command below:

```powershell
Set-ExecutionPolicy `
    -Scope CurrentUser `
    -ExecutionPolicy Bypass
```

[2]: https://www.hancom.com/product/productWindowsMain.do?gnb0=23&gnb1=29

## How to run

When built, the spreadsheet file should contain four macros:

* `Init` macro to initialize the canvas,
* `Render` macro to start rendering a scene,
* `Benchmark` macro to measure how long it takes to render, and
* `Sketch` macro as a "playground" macro.

First, open the built `raytracing.cell` file.

Then, go to `도구`, `매크로`, then `매크로 실행`.
A popup should appear with the list of macros available.
Select and run `Init` macro first, and then follow the same steps to run the
`Render` macro.

With ~~excruciatingly long~~ rendering time, your scene should be available!

Note that due to single-threaded nature of Hancell macros, the rendering process
_will_ block any interaction with the spreadsheet window.
In order to see the rendering process, make sure to put the window on top and
never change focus to another window.
