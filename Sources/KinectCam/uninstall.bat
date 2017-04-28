@D:
@cd D:\KinectV2GreenScreen\

@c:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe /unregister /nologo BaseClassesNET.dll
@c:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe /unregister /nologo KinectV2GreenScreen.dll

@c:\Windows\Microsoft.NET\Framework\v4.0.30319\ngen.exe uninstall BaseClassesNET.dll
@c:\Windows\Microsoft.NET\Framework\v4.0.30319\ngen.exe uninstall KinectV2GreenScreen.dll
