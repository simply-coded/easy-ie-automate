# EasyIEAutomate *(object)*

### An automation wrapper class around the InternetExplorer object that makes it easy to control.   
*( work in progress ) . . .*

##### <p style="text-align:center;color:grey;">[SETUP](#setup) | [USAGE](#usage) </p>

# Setup
> First, let's add the class to our own VBScript file.
### 1. Internet connection available.
   ```vb
    With CreateObject("Msxml2.XMLHttp.6.0")
      Call .open("GET", "https://raw.githubusercontent.com/simply-coded/easy-ie-automate/master/eiea.vbs", False)
	  Call .send() : Call Execute(.responseText)
    End With


    'your code here...
   ```

### 2. No/Unwanted reliance for an internet connection.
  * Download the `eiea.vbs` file.
  * You can **copy & paste** or **import** the code into your VBScript file.  

    * Copy & Paste  
    ```vb           
    Class EasyIEAutomate 
      'paste the EasyIEAutomate class in place of this one.
    End Class


    'your code here...
    ```
    * Import 
    ```vb   
    Dim eieaPath : eieaPath = "c:\path\to\eiea.vbs"
    Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile(eieaPath, 1).ReadAll)    


    'your code here...
    ```

# Usage
> Now that we've added the class to our VBScript file, let's create an instance of it.  

### Create the EasyIEAutomate Object.
```vb
Set IEA = (New EasyIEAutomate)(Null)


'your code here
```
* If you already have a window open that you want to use you can use that in place of the *Null* keyword.
```vb
'URL of the IE window that is already open.
openedIEWindowURL = "https://www.google.com/?gws_rd=ssl"

'Loop through all windows.
Set IEA = Nothing
For Each Window In CreateObject("Shell.Application").Windows
  If Window.LocationURL = openedIEWindowURL Then
    Set IEA = (New EasyIEAutomate)(Window)
    Exit For
  End If
Next

'If window couldn't be found then quit.
If IEA Is Nothing Then
  MsgBox "Couldn't find the Window.",vbCritical 
  WScript.Quit
End If


'IEA reference is now set
'your code here
```

### All the old *properties* and *methods* can be accessed with  the **Base** Property.
```vb
Set IEA = (New EasyIEAutomate)(Null)

'Navigate to Google.
IEA.Base.Navigate "http://www.google.com/"

'Adjust the look of the window.
IEA.Base.AddressBar = False
IEA.Base.MenuBar = False
IEA.Base.StatusBar = False
IEA.Base.Height = 500
IEA.Base.Width = 800

'Wait for window to load
While IEA.Base.Busy : WScript.Sleep(400) : Wend

'Center the window on screen.
Dim SCR : Set SCR = IEA.Base.Document.ParentWindow.screen
IEA.Base.Left = (SCR.width - IEA.Base.Width) / 2
IEA.Base.Top = (SCR.height - IEA.Base.Height) / 2

'Show the window
IEA.Base.Visible = True

'Change the title of the window
IEA.Base.Document.title = "Google Searcher"
```
* All properties and methods can be found [here](https://msdn.microsoft.com/en-us/library/aa752084(v=vs.85).aspx).
