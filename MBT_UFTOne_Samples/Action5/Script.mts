Dim iURL, objShell, fileSystemObj, browserPath, browserName

iURL = "https://www.saucedemo.com/"


Set objShell = CreateObject("Shell.Application")
Set fileSystemObj = CreateObject("Scripting.FileSystemObject")

If fileSystemObj.FileExists("C:\Program Files\Google\Chrome\Application\chrome.exe") Then
    browserPath = "C:\Program Files\Google\Chrome\Application\chrome.exe"   
    browserName = "chrome.exe"
    Else
    Reporter.ReportEvent micFail, "Browser Launch", "No supported browser found"
    ExitTest
End If

objShell.ShellExecute browserPath, iURL, "", "", 1
Wait(5)

if(Browser("Browser").Page("Swag Labs").WebEdit("user-name").Exist(3)) then @@ hightlight id_;_2100766_;_script infofile_;_ZIP::ssf1.xml_;_
   Browser("Browser").Page("Swag Labs").WebEdit("user-name").Set "standard_user" @@ script infofile_;_ZIP::ssf2.xml_;_
    Reporter.ReportEvent micPass, "Username Set", "Username set successfully"
Else
    Reporter.ReportEvent micFail, "Username Not Found", "Failed to find username field"
End If

   
if(Browser("Browser").Page("Swag Labs").WebEdit("password").Exist(3)) then
     Browser("Browser").Page("Swag Labs").WebEdit("password").SetSecure "68934971942fc0f79c0b5fd7f049ec58dd4f58bdb257bb9b775edd951558" @@ script infofile_;_ZIP::ssf3.xml_;_
       Reporter.ReportEvent micPass, "password Set", "password set successfully"
Else
    Reporter.ReportEvent micFail, "password Not Found", "Failed to find password field"
End If

Browser("Browser").Page("Swag Labs").WebButton("Login").Click @@ script infofile_;_ZIP::ssf4.xml_;_

Browser("Browser").Page("Swag Labs").WebButton("Add to cart").Click @@ script infofile_;_ZIP::ssf5.xml_;_

Browser("Browser").Page("Swag Labs").WebButton("Remove").Click

Browser("Browser").Page("Swag Labs").WebButton("Open Menu").Click @@ script infofile_;_ZIP::ssf6.xml_;_

Browser("Browser").Page("Swag Labs").Link("Logout").Click @@ script infofile_;_ZIP::ssf7.xml_;_

SystemUtil.CloseProcessByName browserName




