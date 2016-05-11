// Makes tool windows appear in the MDI document space.

var windows = dte.Windows;

for (var i = 1; i <= windows.Count; i++) {
    var window = windows.Item(i);

    // Check that this is a tool window and not a document
    if (window.Document == null) {
        // Turn off auto-hiding
        if (window.AutoHides) {
            window.AutoHides = false;
        }

        // Set to undockable (which means show the document as maximized)
        try {
            window.Linkable = false;
        } catch (e) { }
    }
}