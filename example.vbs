'Import the EasyIEAutomate Class
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("main.vbs", 1).ReadAll)

'Create an instance of the class
Set IE = (New EasyIEAutomate)(Null)

IE.Visible = True
IE.Navigate "https://rawgit.com/simply-coded/easy-ie-automate/master/practice/index.html"

'Examples of different ways to get an element
    'If it has an ID
    IE.GetElementByID("user").Value = "test 1"
    IE.Base.Document.getElementById("user").Value = "test 2"
    IE.Query("#user").Value = "test 3"
    IE.Base.Document.querySelector("#user").Value = "test 4"

    'If it doesn't have an ID, use css selectors to narrow it down
    'https://www.w3schools.com/cssref/css_selectors.asp
    IE.Query("#form-id > input[type=""password""]").Value = "test 5"

