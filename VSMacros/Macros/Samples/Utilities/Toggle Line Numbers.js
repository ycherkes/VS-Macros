// Toggles line numbering for common languages

var lang = ["Basic", "Plaintext", "CSharp", "HTML", "C/C++", "XML", "TypeScript"];

for (var i = 0; i < lang.length; i++) {
    var currentLang = lang[i];
    try {
        var showLineNumbers = dte.Properties("TextEditor", currentLang).Item("ShowLineNumbers");
        showLineNumbers.Value = !showLineNumbers.Value
    } catch (e) {
        var vsWindowKindOutput = "{34E76E81-EE4A-11D0-AE2E-00A0C90FFFC3}";
        var window = dte.Windows.Item(vsWindowKindOutput);
        var outputWindow = window.Object;
        outputWindow.ActivePane.Activate();
        outputWindow.ActivePane.OutputString("Error occured: " + e.message + ", current language: " + currentLang + "\r\n");
    }
}