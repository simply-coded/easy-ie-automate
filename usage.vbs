'Import the EasyIEAutomate Class
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("eiea.vbs", 1).ReadAll)

'Create an instance of the class
Dim IE : Set IE = (New EasyIEAutomate)(Null)

'Practice HTML page
IE.Navigate "https://rawgit.com/simply-coded/easy-ie-automate/master/practice/index.html"

'Appearance settings
IE.Base.AddressBar = False
IE.Base.MenuBar = False
IE.Visible = True


'Input Data
IE.Query("#user").Value = "Jeremy"
IE.Query("#form-id > input[type=""password""]").Value = "mypassword123"
