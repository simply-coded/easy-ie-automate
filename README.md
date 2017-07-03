# EasyIEAutomate *(object)*
An automation wrapper class around the InternetExplorer object that makes it easy to control with VBScript.

<p align="center">
  <b><a href="#setup">SETUP</a></b> | 
  <b><a href="#usage">USAGE</a></b> | 
  <b><a href="#ie-object">IE-OBJECT</a></b> | 
  <b><a href="#properties">PROPERTIES</a></b> | 
  <b><a href="#methods">METHODS</a></b> | 
  <b><a href="#history">HISTORY</a></b> | 
  <b><a href="#meta">CONTACT</a></b>
</p>

# Setup
First, let's add the class to our own VBScript file.
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

<p align="center">
  <b><a href="#setup">SETUP</a></b> | 
  <b><a href="#usage">USAGE</a></b> | 
  <b><a href="#ie-object">IE-OBJECT</a></b> | 
  <b><a href="#properties">PROPERTIES</a></b> | 
  <b><a href="#methods">METHODS</a></b> | 
  <b><a href="#history">HISTORY</a></b> | 
  <b><a href="#meta">CONTACT</a></b>
</p>

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

<p align="center">
  <b><a href="#setup">SETUP</a></b> | 
  <b><a href="#usage">USAGE</a></b> | 
  <b><a href="#ie-object">IE-OBJECT</a></b> | 
  <b><a href="#properties">PROPERTIES</a></b> | 
  <b><a href="#methods">METHODS</a></b> | 
  <b><a href="#history">HISTORY</a></b> | 
  <b><a href="#meta">CONTACT</a></b>
</p>

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

<p align="center">
  <b><a href="#setup">SETUP</a></b> | 
  <b><a href="#usage">USAGE</a></b> | 
  <b><a href="#ie-object">IE-OBJECT</a></b> | 
  <b><a href="#properties">PROPERTIES</a></b> | 
  <b><a href="#methods">METHODS</a></b> | 
  <b><a href="#history">HISTORY</a></b> | 
  <b><a href="#meta">CONTACT</a></b>
</p>

# Properties
Now let's get into some of the new properties added.

## eIE.Avail 
> **@return**  
[array] - An array of available IE processes (windows and tabs).
```vb
' EXAMPLE 1:
Set eIE = New EasyIEAutomate

' Get number of tabs/windows opened.
count = UBound(eIE.Avail) + 1

' Collect their names & url and show them.
collect = ""
For Each ie In eIE.Avail
  collect = collect & ie.LocationName & " - " & ie.LocationURL & vbLF  
Next

MsgBox collect, vbOKOnly, "IE object(s) open = " & count
```

```vb
' EXAMPLE 2
Set google = New EasyIEAutomate

' Search for a tab/window using google
For Each ie In google.Avail
  If InStr(ie.LocationURL, "google.com/") Then
    google(ie)
    Exit For
  End If
Next

If google.Base Is Nothing Then
  ans = MsgBox("IE with google was not found. Create one?", vbYesNo + vbQuestion)
  If ans = vbYes Then        
    google(vbUseDefault) ' This creates a new IE process
    google.Base.Navigate "https://www.google.com/"    
  Else
    WScript.Quit
  End If
End If

' Show the window if it was hidden or a new one was created.
google.Base.Visible = True

' Wait for it to load before trying to mess with it
While google.Base.Busy : WScript.Sleep 400 : Wend
    
' Search for something in google.
google.Base.Document.getElementById("lst-ib").Value = "searching in google"
google.Base.Document.getElementById("tsf").Submit

WScript.Sleep 2000

' This would be a better way to search, but these are just examples.
google.Base.Navigate "https://www.google.com/#q=alternative+search+in+google"
```

## eIE.Base
> **@return**  
[object] - The main Internet Explorer object. See [IE object](#ie-object).

## eIE.Url
> **@return**  
[string] - The URL of the current IE process. Same as **eIE.Base.LocationURL**. Alerts if no IE process exists.

## eIE.Title
> **@return**  
[string] - The title of the current IE process. Same as **eIE.Base.LocationName**. Alerts if no IE process exists.  

<p align="center">
  <b><a href="#setup">SETUP</a></b> | 
  <b><a href="#usage">USAGE</a></b> | 
  <b><a href="#ie-object">IE-OBJECT</a></b> | 
  <b><a href="#properties">PROPERTIES</a></b> | 
  <b><a href="#methods">METHODS</a></b> | 
  <b><a href="#history">HISTORY</a></b> | 
  <b><a href="#meta">CONTACT</a></b>
</p>

# Methods

## eIE.Close( )
> Closes the current IE process. Same as **eIE.Base.Quit**. Alerts if no IE process exists.

## eIE.CloseAll( )
> Closes all open IE processes (hidden or visible).

## eIE.Show( )
> Sets the visiblilty of the current IE process to true. Same as **eIE.Base.Visible = True**. Alerts if no IE process exists, and then creates one.

## eIE.Hide( )
> Sets the visiblilty of the current IE process to false. Same as **eIE.Base.Visible = False**. Alerts if no IE process exists, and then creates one.

## eIE.Center( )
> Centers the window on screen. If no website is loaded one with "about:blank" will be navigated to. A loaded website is necessary to get the screen resolution. 
```vb
Set eIE = (New EasyIEAutomate)(vbUseDefault)

' Centering before the browser is visible makes for a nicer looking experience.
eIE.Center
eIE.Show
```

## eIE.Navigate( url )
> Navigates the IE process to this location. Same as **eIE.Base.Navigate2 "URL_HERE"**. Alerts if no IE process exists, and then creates one.  
> **@params**   
*url* [string] - An address url to navigate to.

## eIE.NavigateTab( url )
> Creates a new tab and navigates the IE process to this location. Same as **eIE.Base.Navigate2 "URL_HERE", 2048**. Alerts if no IE process exists, and then creates one.  
> **@params**   
*url* [string] - An address url to navigate to.

## eIE.NavigateBgTab( url )
> Creates a new background tab and navigates the IE process to this location. Same as **eIE.Base.Navigate2 "URL_HERE", 4096**. Alerts if no IE process exists, and then creates one.  
> **@params**   
*url* [string] - An address url to navigate to. 

## eIE.WaitForLoad( )
> Waits for the current IE process to finish loading. Alerts if no IE process exists.

## eIE.DeepWaitForLoad( elem )
> The method waits for the inputed element to finish loading. The current page could be loaded but content inside of it like an iframe, could still need time to load.  
> **@params**  
*elem* [object] - An HTML element object like an iframe. 
```vb
Set eIE = (New EasyIEAutomate)(vbUseDefault)

eIE.Navigate "https://rawgit.com/simply-coded/easy-ie-automate/master/practice/index.html"
eIE.Show

' Wait for main webpage to load
eIE.WaitForLoad

' Get an iframe on that page. 
Set myFrame = eIE.Base.Document.getElementById("ice_frame")

' Wait for iframe to load if it hasn't already.
eIE.DeepWaitForLoad(myFrame)

' To access things inside an iframe it must respect the same origin policy!
MsgBox myFrame.contentDocument.querySelector("img").src, vbOKOnly, "The image source is:"
```
There's an easier way however. The next example does the same thing as this one by using a function called **Deeper()**.

## eIE.Deeper( squery )
> Method waits for element to load and then returns the frame's contentDocument. Alerts and quits the script if same origin policy is violated. Used for accessing/editing things within a frame.  
> **@params**  
*squery* [string] - A query string to get the iframe.  
**@return**  
[object] - The frame element's contentDocument as an object.
```vb
Set eIE = (New EasyIEAutomate)(vbUseDefault)

eIE.Navigate "https://rawgit.com/simply-coded/easy-ie-automate/master/practice/index.html"
eIE.Show

' Waits for main webpage to load, selects the element, waits for it to load, 
' and then returns the element's contentDocument. Alerts and quits if same 
' origin policy is violated.
Set myFrame = eIE.Deeper("#ice_frame")
MsgBox myFrame.querySelector("img").src, vbOKOnly, "The image source is:"
```

## eIE.ReBase( ie )
> Changes the Base to a different Internet Explorer object. Alias for the default Init() method. An example can be found in the [usage](#usage) section.  
**@params**  
*ie* [object] - The Internet Explorer window/tab object you want to change the Base object to. 

## eIE.RePoint( url )
> Changes the Base to a different Internet Explorer object. An example can be found in the [usage](#usage) section.  
**@params**    
*url* [string] - An url of an Internet Explorer window/tab that is currently opened that you want to change the Base object to. 

## eIE.Latest( )
> Changes the Base to the most recent Internet Explorer object. An example can be found in the [usage](#usage) section.  

## eIE.Query( squery )
> Uses *eIE.Base.Document.querySelector( squery )* to retrieve an element. [Here](https://developer.mozilla.org/en-US/docs/Web/API/Element/querySelector) is documentation on Element.querySelector(). [Here](https://developer.mozilla.org/en-US/docs/Web/CSS/CSS_Selectors) are the various ways to search for elements. Waits for document to load before searching, and alerts and quits if it cannot be found.   
**@params**    
*squery* [string] - The query string to select the element with.   
**@return**  
[object] - The found element as an object. 
```vb
Set eIE = (New EasyIEAutomate)(vbUseDefault)

eIE.Navigate "https://rawgit.com/simply-coded/easy-ie-automate/master/practice/index.html"
eIE.Show

' get various elements and interact with them
eIE.Query(".titles").style.color = "deepskyblue"
eIE.Query("input[name='email']").Value = "simplycoded.help@gmail.com"
eIE.Query("#pass").Value = "bananas_are_the_universal_scale"
eIE.Query("#milk").removeAttribute("checked")
eIE.Query("input[type='radio'][value='female']").setAttribute("checked")
eIE.Query("form > p > button").Click
```

## eIE.QueryAll( squery )
> Uses *eIE.Base.Document.querySelectorAll( squery )* to retrieve a NodeList. [Here](https://developer.mozilla.org/en-US/docs/Web/CSS/CSS_Selectors) are the various ways to search for elements. Waits for document to load before searching, and alerts and quits if something cannot be found.   
**@params**    
*squery* [string] - The query string to select the elements with.   
**@return**  
[object] - The found elements as a NodeList object. 
```vb
Set eIE = (New EasyIEAutomate)(vbUseDefault)
eIE.Navigate "https://rawgit.com/simply-coded/easy-ie-automate/master/practice/index.html"
eIE.Show

' get elements with class="titles" and all <p> tag elements
Set list = eIE.QueryAll(".titles, p")

' change their style up a bit
For i = 0 To list.length - 1
  list.item(i).style.backgroundColor = "white"  
  list.item(i).style.border = "2px solid deepskyblue"  
  list.item(i).style.fontFamily = "Consolas, monospace"
Next
```

<p align="center">
  <b><a href="#setup">SETUP</a></b> | 
  <b><a href="#usage">USAGE</a></b> | 
  <b><a href="#ie-object">IE-OBJECT</a></b> | 
  <b><a href="#properties">PROPERTIES</a></b> | 
  <b><a href="#methods">METHODS</a></b> | 
  <b><a href="#history">HISTORY</a></b> | 
  <b><a href="#meta">CONTACT</a></b>
</p>

# History
* 1.1
  * `ADD` - Documentation added to *README.md*.
  * `CHANGE` - Methods renamed as follows: **NavigateT()** to **NavigateTab()**, and **NavigateBT()** to **NavigateBgTab()**.
  * `CHANGE` - **ReBase()** now an alias for Default **Init()** method.
  * `ADD` - Timed popups for errors not worth quitting over. See documentation for what methods alert.
  * `CHANGE` - Practice folder examples have changed in *challenge.vbs* and *answers.vbs* to reflect newest syntax.
* 1.0
  * `ADD` - Initial release.

<p align="center">
  <b><a href="#setup">SETUP</a></b> | 
  <b><a href="#usage">USAGE</a></b> | 
  <b><a href="#ie-object">IE-OBJECT</a></b> | 
  <b><a href="#properties">PROPERTIES</a></b> | 
  <b><a href="#methods">METHODS</a></b> | 
  <b><a href="#history">HISTORY</a></b> | 
  <b><a href="#meta">CONTACT</a></b>
</p>

# Meta
Jeremy England ( SimplyCoded ) - [simplycoded.help@gmail.com](mailto:simplycoded.help@gmail.com)

Distributed under the MIT license. See [`LICENSE`](https://spdx.org/licenses/MIT.html) for more information.
