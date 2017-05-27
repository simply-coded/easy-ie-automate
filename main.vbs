'''
' @author Jeremy England
' @company SimplyCoded
' @date 01/26/2017
' @description
'  Makes automating Internet Explorer easier.
' @tutorial
'  http://link.here/
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
	  End Sub 
    
    Private Sub Class_Terminate
        Set classIE = Nothing
        Set classSHELL = Nothing
    End Sub

    Public Default Function Init(objIE)
        If isnull(objIE) Then
            Set classIE = CreateObject("InternetExplorer.Application")
        Else 
            Set classIE = objIE
        End If
        Set Init = Me
    End Function 
    '''
    ' PROPERTIES
    '
    Public Property Get Base
        Set Base = classIE 
    End Property   
    
    Public Property Get URL
        URL = classIE.LocationURL
    End Property
    
    Public Property Get Title
        Title = classIE.LocationName
    End Property    
        
    Public Property Get Tabs
        Call Wait()
        Dim window, list
        list = Array()
        For Each window In classSHELL.Windows
            If InStr(LCase(window.FullName), "iexplore.exe") Then 
                ReDim Preserve list(UBound(list) + 1)
                Set list(UBound(list)) = (New EasyIEAutomate)(window)
            End If
        Next
        Tabs = list
	  End Property

    Public Property Let Visible(bool)
        classIE.Visible = bool
    End Property
    '''
    ' SUBS
    '    
    Public Sub Close()
        classIE.Quit()
    End Sub

    Public Sub CloseAll()
        Dim window
        For Each window In classSHELL.Windows
            If InStr(LCase(window.FullName), "iexplore.exe") Then                                      
                window.Quit()
            End If
        Next 
    End Sub

    Public Sub Wait()
        While (classIE.Busy) And Not (classIE.ReadyState = 4) : WScript.Sleep(400) : Wend 
    End Sub

    Public Sub RePoint(strURL)  
        Dim tab      
        For Each tab in Tabs()
            tab.Wait()
            If (LCase(tab.URL) = LCase(strURL)) Then
                Set classIE = tab.Base
                Exit Sub
            End If
        Next
        Call ErrorOut(Array("Internet Explorer", strURL))        
    End Sub

    Public Sub Latest()
        Call Wait()
        Dim window
        For Each window In classSHELL.Windows
            If InStr(LCase(window.FullName), "iexplore.exe") Then                                       
                Set classIE = window                
            End If
        Next                
	  End Sub
    '''
    ' FUNCTIONS
    '
    Public Function Navigate(strURL)
        classIE.Navigate2 strURL
    End Function
    
    Public Function NavigateT(strURL)
        classIE.Navigate2 strURL, 2048
    End Function

    Public Function NavigateBT(strURL)
        classIE.Navigate2 strURL, 4096
    End Function
    
    Public Function GetElementByTagText(strTag, strText)
        Call Wait()
        Dim tag
        For Each tag In classIE.Document.GetElementsByTagName(strTag)
            If InStr(LCase(tag.innerText),  LCase(strText)) Then
                Set GetElementByTagText = tag   
                Exit Function
            End If 
        Next
        Call ErrorOut(Array(classIE.LocationURL, strTag, strText))        
    End Function
    
    Public Function GetElementByID(strID)
        Call Wait()
        On Error Resume Next
        Dim element
        Set element = classIE.Document.getElementByID(strID)
        If Err.Number = 0 Then 
            Set GetElementByID = element
        Else
            Call ErrorOut(Array(classIE.LocationURL, strID))        
        End If
    End Function
    
    Public Function Query(strSelector)
        Call Wait()
        On Error Resume Next
        Dim element
        Set element = classIE.Document.querySelector(strSelector)
        If Err.Number = 0 Then
            Set Query = element
        Else
            Call ErrorOut(Array(classIE.LocationURL, strSelector))
        End If 
    End Function 
    '''
    ' ERROR HANDLING
    '    
    Private Sub ErrorOut(args)        
        If (UBound(args) = 1) Then
            MsgBox _
                "CANNOT FIND [ " & args(1) & " ]" & vbLf & _
                "AT [ " & args(0) & " ]", vbCritical
        ElseIf (UBound(args) = 2) Then
            MsgBox _
                "CANNOT FIND [ " & args(1) & " ]" & vbLf & _
                "WITH [ " & args(2) & " ]" & vbLf & _
                "AT [ " & args(0) & " ]", vbCritical
        End If                
        WScript.quit
    End Sub
End Class