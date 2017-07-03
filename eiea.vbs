Class EasyIEAutomate
  '''
  ' OBJECTS
  ' 
  Private classIE
  Private classWSS  
  Private classSHELL  
 
  '''
  ' EVENTS
  ' 
  Private Sub Class_Initialize
    Set classIE = Nothing
    Set classWSS = CreateObject("WScript.Shell")
    Set classSHELL = CreateObject("Shell.Application")     
  End Sub 
  
  Private Sub Class_Terminate
    Set classIE = Nothing
    Set classWSS = Nothing
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
  ' METHODS
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
  
  Public Sub Center()        
    Call WaitForLoad()   
    On Error Resume Next  
    With classIE.Document.ParentWindow.screen
      classIE.Left = (.width - classIE.Width) / 2
      classIE.Top = (.height - classIE.Height) / 2
    End With    
    If Err.Number = 505 Then 
       Navigate "about:blank"
       Center()
    End If
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
    
  Public Sub ReBase(ie)
    Init(ie)
  End Sub
  
  Public Sub RePoint(url)  
    Dim window
    For Each window in classSHELL.Windows            
      If IsIE(window) And (LCase(window.LocationURL) = LCase(url)) Then             
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
    
  Private Function IsIE(obj)    
    IsIE = CBool(Right(LCase(obj.FullName), 12) = "iexplore.exe")
  End Function
  
  Public Function Query(squery)
    Call WaitForLoad()
    On Error Resume Next
    Dim element
    Set element = classIE.Document.querySelector(squery)
    If Err.Number = 0 Then
      Set Query = element
    Else
      Call ErrorOut(squery, classIE.LocationURL)
    End If 
  End Function
  
  Public Function QueryAll(squery)
    Call WaitForLoad()
    On Error Resume Next
    Dim elements
    Set elements = classIE.Document.querySelectorAll(squery)
    If Err.Number = 0 Then
      Set QueryAll = elements
    Else
      Call ErrorOut(squery, classIE.LocationURL)
    End If 
  End Function
  
  Public Function Deeper(squery)
    Call WaitForLoad()
    On Error Resume Next
    Dim element
    Set element = classIE.Document.querySelector(squery)
    If Err.Number = 0 Then
      Call DeepWaitForLoad(element)
      Set Deeper = element.contentDocument
      If Err.Number = -2147024891 Then
        MsgBox "ERROR: Deeper(""" & squery & """)" & vbLf & vbLf & "Same Origin Policy Violated.", vbCritical, "EasyIEAutomate: " & Err.Description
        WScript.Quit
      End If
    Else
      Call ErrorOut(squery, classIE.LocationURL)
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
      classWSS.PopUp "Auto initialized a new IE object.", 1, "EasyIEAutomate"      
      Set classIE = CreateObject("InternetExplorer.Application")
    End If 
  End Sub 
  
  Private Sub noInitMsg()    
    classWSS.PopUp "Not yet initialized.", 1, "EasyIEAutomate"    
  End Sub
  
End Class
