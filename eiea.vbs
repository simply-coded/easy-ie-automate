'''
' @author Jeremy England
' @company SimplyCoded
' @date 06/24/2017
' @description
'  Makes automating Internet Explorer easier.
'

'''
' INTERNET EXPLORER AUTOMATION CLASS
'
Class EasyIEAutomate
  '''
  ' OBJECTS
  ' 
  Private classIE
  Private classSHELL    
 
  '''
  ' EVENTS
  ' 
  Private Sub Class_Initialize
    Set classSHELL= CreateObject("Shell.Application") 
    Set classIE = Nothing
  End Sub 
  
  Private Sub Class_Terminate
    Set classIE = Nothing
    Set classSHELL = Nothing
  End Sub
  
  '''
  ' CONSTRUCTOR
  '  
  Public Default Function Init(obj)  
    If IsObject(obj) Then
      If obj Is Nothing Then
        Set classIE = Nothing
      ElseIf IsIE(obj) Then 
        Set classIE = obj
      End If
    ElseIf IsNumeric(obj) Then
      If obj = vbUseDefault Then 
        Set classIE = CreateObject("InternetExplorer.Application")
      End If    
    End If  
    Set Init = Me
  End Function 
     
  '''
  ' PROPERTIES
  '  
  Public Property Get Avail
    Dim process, list : list = Array()
    For Each process In classSHELL.Windows
      If IsIE(process) Then 
        ReDim Preserve list(UBound(list) + 1)
        Set list(UBound(list)) = process
      End If
    Next
    Avail = list
  End Property 
  
  Public Property Get Base    
    If classIE Is Nothing Then 
      Call noInitMsg()
    End If
    Set Base = classIE
  End Property   
  
  Public Property Get Url
    If classIE Is Nothing Then
      Call noInitMsg()
      Url = ""      
    Else
      Url = classIE.LocationURL
    End If
  End Property
  
  Public Property Get Title
    If classIE Is Nothing Then
      Call noInitMsg()
      Title = ""
    Else
      Title = classIE.LocationName
    End If    
  End Property
  
  '''
  ' SUBS
  '       
  Public Sub Close()
    If classIE Is Nothing Then 
      Call noInitMsg()      
    Else
      classIE.Quit
    End If      
  End Sub
  
  Public Sub CloseAll()
    Dim window
    For Each window In classSHELL.Windows
      If IsIE(window) Then : window.Quit : End If
    Next 
  End Sub
  
  Public Sub Show()
    Call autoInit()
    classIE.Visible = True
  End Sub
  
  Public Sub Hide()
    Call autoInit()
    classIE.Visible = False
  End Sub
  
  Public Sub Navigate(url)
    Call autoInit()
    classIE.Navigate2 url
  End Sub
  
  Public Sub NavigateTab(url)
    Call autoInit()
    classIE.Navigate2 url, 2048
  End Sub
  
  Public Sub NavigateBgTab(url)
    Call autoInit()
    classIE.Navigate2 url, 4096
  End Sub
  
  Public Sub WaitForLoad()
    If classIE Is Nothing Then 
      Call noInitMsg()      
    Else
      While (classIE.Busy) And Not (classIE.ReadyState = 4) : WScript.Sleep(400) : Wend   
    End If    
  End Sub
  
  Public Sub DeepWaitForLoad(elem)                
    While Not (elem.ReadyState = "complete") : WScript.Sleep(400) : Wend        
  End Sub
  
  
  Public Sub ReBase(obj)
    Init(obj)
  End Sub
  
  Public Sub RePoint(strURL)  
    Dim window
    For Each window in classSHELL.Windows            
      If IsIE(window) And (LCase(window.LocationURL) = LCase(strURL)) Then             
        Set classIE = window
        Exit Sub
      End If
    Next
    Call ErrorOut(strURL, "Internet Explorer")
  End Sub
  
  Public Sub Latest()        
    Dim window
    For Each window In classSHELL.Windows
      If IsIE(window) Then
        Set classIE = window 
        Exit Sub
      End If
    Next         
    Call autoInit()
  End Sub
  
  '''
  ' FUNCTIONS
  '  
  Private Function IsIE(obj)    
    IsIE = CBool(Right(LCase(obj.FullName), 12) = "iexplore.exe")
  End Function
  
  Public Function Query(strSelector)
    Call WaitForLoad()
    On Error Resume Next
    Dim element
    Set element = classIE.Document.querySelector(strSelector)
    If Err.Number = 0 Then
      Set Query = element
    Else
      Call ErrorOut(strSelector, classIE.LocationURL)
    End If 
  End Function
  
  Public Function Deeper(strSelector)
    Call WaitForLoad()
    On Error Resume Next
    Dim element
    Set element = classIE.Document.querySelector(strSelector).contentDocument
    If Err.Number = 0 Then
      Call DeepWaitForLoad(element)
      Set Deeper = element.documentElement
    Else
      Call ErrorOut(strSelector, classIE.LocationURL)
    End If 
  End Function   
  
  '''
  ' ERROR HANDLING
  '      
  Private Sub ErrorOut(item, at)        
    MsgBox _
      "CANNOT FIND [ " & item & " ]" & vbLf & _
      "AT [ " & at & " ]", vbCritical
    WScript.quit
  End Sub
  
  Private Sub autoInit()
    If classIE Is Nothing Then
      With CreateObject("WScript.Shell")
        .PopUp "Auto initialized a new IE object.", 1, "EasyIEAutomate"
      End With
      Set classIE = CreateObject("InternetExplorer.Application")
    End If 
  End Sub 
  
  Private Sub noInitMsg()    
    With CreateObject("WScript.Shell")
      .PopUp "Not yet initialized.", 1, "EasyIEAutomate"
    End With    
  End Sub
  
End Class

' EXAMPLE 1:
Dim eIE
Set eIE = New EasyIEAutomate

eIE.ReBase eIE.Avail()(0)

Set wFrame = eIE.Base.Document.getElementById("weather")
'wFrame.addEventListener "load", GetRef("deepwait")

MsgBox wFrame.contentWindow.document.querySelector("#current_conditions-summary > .myforecast-current-sm").innerText