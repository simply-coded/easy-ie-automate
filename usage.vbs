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
Set iframe = IE.Query("#ice_frame").contentDocument.documentElement

Set selectElem = iframe.querySelector("select")
For i = 0 To selectElem.options.length - 1
    If selectElem.options(i).value = "mint" Then
        selectElem.options(i).selected = true
    Else
        selectElem.options(i).selected = false
    End If
Next




'----(5) Task--------------------------------
        'Select Menu Dropdown Option - [Tell Us More]
        'Input Text - [Tell us about you.]
        'Click - [Submit >]
'--------------------------------------------



