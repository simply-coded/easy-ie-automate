'Import the EasyIEAutomate Class
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("eiea.vbs", 1).ReadAll)

'Create an instance of the class
Dim IE : Set IE = (New EasyIEAutomate)(Null)

'Practice HTML page
IE.Navigate "https://rawgit.com/simply-coded/easy-ie-automate/master/practice/index.html"
IE.Base.Visible = True

'Challenges: One of multiple ways to do this can be found in answers.vbs

'----(1) Task--------------------------------
        'Input Text - [Name, Email, Password]
'--------------------------------------------




'----(2) Task--------------------------------
        'Uncheck - [Milk]
        'Check - [Sugar, Lemon]
        'Select - [Female]
'--------------------------------------------




'----(3) Task--------------------------------
        'Click - [Send >]
'--------------------------------------------




'----(4) Task--------------------------------
        'Select Icecream Dropdown Option - [Mint]
        'Click - [Save >]
'--------------------------------------------




'----(5) Task--------------------------------
        'Select Menu Dropdown Option - [Tell Us More]
        'Input Text - [Tell us about you.]
        'Click - [Submit >]
'--------------------------------------------



