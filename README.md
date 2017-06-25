# EasyIEAutomate *(object)*

### An automation wrapper class around the InternetExplorer object that makes it easy to control.   
*( work in progress ) . . .*

##### <p style="text-align:center;color:grey;">[SETUP](#setup) | [USAGE](#usage) | [IE OBJECT](#ie-object) </p>

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

## Initialization 
```vb
Set eIE = New EasyIEAutomate

'your code here
```
> By default this will **not** create an IE process immediately. The reason for this is so that if you have an existing IE window or tab open you can grab that process instead of creating a new one. Creating or getting an existing window can be achieved in a few ways: **Init()**, **ReBase()**, **RePoint()**, and **Latest()**.  

> Let's see some examples:

* Create new with **Init()**.
```vb
' These all do the same thing.
'1.
Set eIE = (New EasyIEAutomate)(vbUseDefault) 

'2.
Set eIE = (New EasyIEAutomate)(CreateObject("InternetExplorer.Application"))

'3.
Set eIE = New EasyIEAutomate
eIE(vbUseDefault)

'4.
Set eIE = New EasyIEAutomate
eIE(CreateObject("InternetExplorer.Application"))
```

* Get existing with **Init()**. 
```vb
'Already have the object
Set objIE = CreateObject("InternetExplorer.Application")
Set eIE = (New EasyIEAutomate)(objIE)
```

* Get existing with **ReBase()**. 
```vb
'Already have the object
Set objIE = CreateObject("InternetExplorer.Application")
Set eIE = New EasyIEAutomate
eIE.ReBase objIE
```
> The method **ReBase()** can be used at anytime to change what IE process the class is controlling and is just an alias for the **Init()** method.

* Get existing with **RePoint()**.
```vb
Set eIE = New EasyIEAutomate 

'URL of the IE window that is already open.
eIE.RePoint "https://www.google.com/?gws_rd=ssl"
```

* Get existing with **Latest()**.
```vb
Set eIE = New EasyIEAutomate 

'Will get the most recent IE process created.
eIE.Latest 
```

## Auto IE Initialization
> If you start off with no IE process EasyIEAutomate will automatically create one when you start using the class. A one (1) second Popup of *"Auto initialized a new IE object."* will let you know if it does.
```vb
Set eIE = (New EasyIEAutomate)(Nothing)
' Or just: Set eIE = New EasyIEAutomate 

' This and many other methods and properties will trigger an automatic creation of an IE process if none exist.
eIE.Show 
```

# IE Object

> All the IE *properties* and *methods* you're use to can still be accessed through the **Base** Property.
```vb
Set IEA = (New EasyIEAutomate)(vbUseDefault)

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
> If all of this looks unfamiliar to you then I would recommend checking out all the main properties and methods [here](https://msdn.microsoft.com/en-us/library/aa752084(v=vs.85).aspx).

# Properties
> Now let's get into some of the new properties added.

### Avail 
@return - an array of available IE processes (windows and tabs).
```vb
Set google = New EasyIEAutomate

For Each ie In google.Avail
  If InStr(ie.LocationURL, "//www.google.com/") Then
    google(ie)  
  End If
Next

If google.Base Is Nothing Then
  ans = MsgBox("IE with Google not found. Create one?", vbYesNo + vbQuestion)
  If ans = vbYes Then
    google.Navigate "https://www.google.com/"        
  Else
    WScript.Quit
  End If
End If

google.Base.Visible = True
```