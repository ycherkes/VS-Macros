// Toggles the word wrapping in the text editor view

var lang = ["Basic", "Plaintext", "CSharp", "HTML", "C/C++", "XML", "TypeScript"];

for (var i = 0; i < lang.length; i++) {
    var currentLang = lang[i];
    try {
        var wordWrap = dte.Properties("TextEditor", lang[i]).Item("WordWrap");
        wordWrap.Value = !wordWrap.Value;
    } catch (e) {
        var vsWindowKindOutput = "{34E76E81-EE4A-11D0-AE2E-00A0C90FFFC3}";
        var window = dte.Windows.Item(vsWindowKindOutput);
        var outputWindow = window.Object;
        outputWindow.ActivePane.Activate();
        outputWindow.ActivePane.OutputString("Error occured: " + e.message + ", current language: " + currentLang + "\r\n");
    }
}