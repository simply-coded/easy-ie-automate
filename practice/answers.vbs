'Import the EasyIEAutomate Class
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("eiea.vbs", 1).ReadAll)

'Create an instance of the class
Dim eIE : Set eIE = (New EasyIEAutomate)(vbUseDefault)

'Practice HTML page
eIE.Navigate "https://rawgit.com/simply-coded/easy-ie-automate/master/practice/index.html"
eIE.Show

'Input in data
'----(1) Task--------------------------------
eIE.Query("#user").Value = "Jeremy"
eIE.Query("input[name='email']").Value = "simplycoded.help@gmail.com"
eIE.Query("#pass").Value = "bananas_are_the_universal_scale"

'----(2) Task--------------------------------
eIE.Query("#milk").removeAttribute("checked")
eIE.Query("#sugar").setAttribute("checked")
eIE.Query("#lemon").setAttribute("checked")
eIE.Query("input[type='radio'][value='female']").setAttribute("checked")

'----(3) Task--------------------------------
eIE.Query("form > p > button").Click

'----(4) Task--------------------------------
Set iframe = eIE.Deeper("#ice_frame")

Set selectElement = iframe.querySelector("select")
For i = 0 To selectElement.options.length - 1
    If selectElement.options(i).value = "mint" Then
        selectElement.options(i).selected = true
    Else
        selectElement.options(i).selected = false
    End If
Next

iframe.querySelector("input[type='submit']").Click

'----(5) Task--------------------------------
eIE.Query("a[href='./more.html']").Click
eIE.RePoint("https://rawgit.com/simply-coded/easy-ie-automate/master/practice/more.html")
eIE.Query("textarea").Value = "I am allergic to butterflies."
eIE.Query("form input").Click
