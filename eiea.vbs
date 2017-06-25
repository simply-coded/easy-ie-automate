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
      If (Right(LCase(process.FullName), 12) = "iexplore.exe") Then 
        ReDim Preserve list(UBound(list) + 1)
        Set list(UBound(list)) = process
      End If
    Next
    Avail = list
  End Property 
  
  Public Property Get Base    
    Set Base = classIE 
  End Property   
  
  Public Property Get URL
    Call checkInit()
    URL = classIE.LocationURL
  End Property
  
  Public Property Get Title
    Call checkInit()
    Title = classIE.LocationName
  End Property
  
  '''
  ' SUBS
  '       
  Public Sub Close()
    Call checkInit()
    classIE.Quit
  End Sub
  
  Public Sub CloseAll()
    Dim window
    For Each window In classSHELL.Windows
      If IsIE(window) Then : window.Quit : End If
    Next 
  End Sub
  
  Public Sub Show()
    Call checkInit()
    classIE.Visible = True
  End Sub
  
  Public Sub Hide()
    Call checkInit()
    classIE.Visible = False
  End Sub
  
  Public Sub Navigate(strURL)
    Call checkInit()
    classIE.Navigate2 strURL
  End Sub
  
  Public Sub NavigateTab(strURL)
    Call checkInit()
    classIE.Navigate2 strURL, 2048
  End Sub
  
  Public Sub NavigateBgTab(strURL)
    Call checkInit()
    classIE.Navigate2 strURL, 4096
  End Sub
  
  Public Sub WaitForLoad()
    Call checkInit()
    While (classIE.Busy) And Not (classIE.ReadyState = 4) : WScript.Sleep(400) : Wend 
  End Sub
  
  Public Sub DeepWaitForLoad(frame)                
    While Not (frame.ReadyState = "complete") : WScript.Sleep(400) : Wend        
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
    Call checkInit()
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
  
  Private Sub checkInit()
    If classIE Is Nothing Then
      With CreateObject("WScript.Shell")
        .PopUp "Auto initialized a new IE object.", 1, "EasyIEAutomate"
      End With
      Set classIE = CreateObject("InternetExplorer.Application")
    End If 
  End Sub 
  
End Class