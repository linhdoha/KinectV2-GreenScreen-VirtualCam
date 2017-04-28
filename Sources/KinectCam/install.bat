@D:
@cd D:\KinectV2GreenScreen\

@c:\Windows\Microsoft.NET\Framework\v4.0.30319\ngen.exe install BaseClassesNET.dll
@c:\Windows\Microsoft.NET\Framework\v4.0.30319\ngen.exe install KinectV2GreenScreen.dll

@c:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe /nologo BaseClassesNET.dll
@c:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe KinectV2GreenScreen.dll /nologo /codebase /tlb: KinectV2GreenScreen.tlb
