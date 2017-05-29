'Import the EasyIEAutomate Class
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("eiea.vbs", 1).ReadAll)

'Create an instance of the class
Dim IE : Set IE = (New EasyIEAutomate)(Null)

'Practice HTML page
IE.Navigate "https://rawgit.com/simply-coded/easy-ie-automate/master/practice/index.html"

'Appearance settings
IE.Base.AddressBar = False
IE.Base.MenuBar = False
IE.Base.Visible = True

'Input in data
'----(1) Task--------------------------------
IE.Query("#user").Value = "Jeremy"
IE.Query("input[name='email']").Value = "simplycoded.help@gmail.com"
IE.Query("#pass").Value = "bananas1are2the3universal4scale5"

'----(2) Task--------------------------------
IE.Query("#milk").removeAttribute("checked")
IE.Query("#sugar").setAttribute("checked")
IE.Query("#lemon").setAttribute("checked")
IE.Query("input[type='radio'][value='female']").setAttribute("checked")

'----(3) Task--------------------------------
IE.Query("form > p > button").Click

'----(4) Task--------------------------------
Set iframe = IE.Deeper("#ice_frame")

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
IE.Query("a[href='./more.html']").Click
IE.RePoint("https://rawgit.com/simply-coded/easy-ie-automate/master/practice/more.html")
IE.Query("textarea").Value = "I am allergic to butterflies."
IE.Query("form input").Click
