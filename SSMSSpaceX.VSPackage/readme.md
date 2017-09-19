C:\Program Files (x86)\Microsoft SQL Server\130\Tools\Binn\ManagementStudio\Extensions\MyAddin
C:\ProgramData\Microsoft\MSEnvShared\Addins

===== how to debug 
1. start external program
C:\Program Files (x86)\Microsoft SQL Server\130\Tools\Binn\ManagementStudio\Ssms.exe
D:\Program Files (x86)\Microsoft Visual Studio\2017\Community\Common7\IDE\devenv.exe /rootsuffix Exp

2. Disable PInvokeStackImbalance exceptions

========
资料? https://sqljudo.wordpress.com/31-days-of-sql-server-management-studio/ssms-day-30-vspackage-and-ssms/

Extensibility in Visual Studio
https://msdn.microsoft.com/zh-CN/library/bb166030.aspx

多版本差异化解决方案例子
http://ssmsschemafolders.codeplex.com/SourceControl/latest#README.md

========


 regedit
 HKEY_CURRENT_USER\SOFTWARE\Microsoft\SQL Server Management Studio\13.0_Config\Packages\{3b76e61c-fe19-493b-b881-23b5e4fc30a5}





