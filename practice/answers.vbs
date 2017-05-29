'Import the EasyIEAutomate Class
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("eiea.vbs", 1).ReadAll)

'Create an instance of the class
Dim IEA : Set IEA = (New EasyIEAutomate)(Null)

'Practice HTML page
IEA.Navigate "https://rawgit.com/simply-coded/easy-ie-automate/master/practice/index.html"
IEA.Base.Visible = True

'Input in data
'----(1) Task--------------------------------
IEA.Query("#user").Value = "Jeremy"
IEA.Query("input[name='email']").Value = "simplycoded.help@gmail.com"
IEA.Query("#pass").Value = "bananas1are2the3universal4scale5"

'----(2) Task--------------------------------
IEA.Query("#milk").removeAttribute("checked")
IEA.Query("#sugar").setAttribute("checked")
IEA.Query("#lemon").setAttribute("checked")
IEA.Query("input[type='radio'][value='female']").setAttribute("checked")

'----(3) Task--------------------------------
IEA.Query("form > p > button").Click

'----(4) Task--------------------------------
Set iframe = IEA.Deeper("#ice_frame")

With iframe.querySelector("select")
    For i = 0 To .options.length - 1
        If .options(i).value = "mint" Then
            .options(i).selected = true
        Else
            .options(i).selected = false
        End If
    Next
End With

iframe.querySelector("input[type='submit']").Click

'----(5) Task--------------------------------
IEA.Query("a[href='./more.html']").Click
IEA.RePoint("https://rawgit.com/simply-coded/easy-ie-automate/master/practice/more.html")
IEA.Query("textarea").Value = "I am allergic to butterflies."
IEA.Query("form input").Click
