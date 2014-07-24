//-----------------------------------------------------------------------
// <summary>
//      Defines an object dte used to provide context to the 
//      javascript intellisense.
// </summary>
// <copyright file="dte.js" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

/// Helper variables

var _dte2 = new function () {
    
    this.Selection = new function () {
        /// <signature>
        /// <summary>Provides access to view-based editing operations and selected text.</summary>
        /// </signature>

        /// Gets the current endpoint of the selection.
        this.ActivePoint = "";

        /// Microsoft Internal Use Only.
        this.AnchorColumn = "";

        /// Gets the origin point of the selection.
        this.AnchorPoint = "";

        /// Microsoft Internal Use Only.
        this.BottomLine = "";

        /// Gets the point at the end of the selection.
        this.BottomPoint = "";

        /// Microsoft Internal Use Only.
        this.CurrentColumn = "";

        /// Microsoft Internal Use Only.
        this.CurrentLine = "";

        /// Gets the top-level extensibility object.
        this.DTE = "";

        /// Gets whether the active pois equal to the bottom po.
        this.IsActiveEndGreater = "";

        /// Gets whether the anchor pois equal to the active po.
        this.IsEmpty = "";

        /// Sets or gets a value determining whether dragging the mouse selects in stream
        this.Mode = "";

        /// Gets the immediate parent object of a EnvDTE.TextSelection object.
        this.Parent = "";

        /// Sets or gets the selected text.
        this.Text = "";

        /// Gets the text pane that contains the selected text.
        this.TextPane = "";

        /// Gets a EnvDTE.TextRanges collection with one EnvDTE.TextRange object for each
        this.TextRanges = "";

        /// Microsoft Internal Use Only.
        this.TopLine = "";

        /// Gets the top end of the selection.
        this.TopPoint = "";

        /// <summary>Microsoft Internal Use Only.</summary>
        /// <param name="Count">The number of spaces.</param>
        this.Backspace = function (Count) { }

        this.ChangeCase = function (How) {
            /// <summary>Microsoft Internal Use Only.</summary>
        }

        this.CharLeft = function (Extend, Count) {
            /// <summary>Moves the object the specified number of characters to the left.</summary>
            /// <param name="Extend">Optional. Determines whether the moved text is collapsed or not. The defaultis False.</param>
            /// <param name="Count">Optional. Represents the number of characters to move to the left. The defaultis 1.</param>
        }

        this.CharRight = function (Extend, Count) {
            /// <summary>
            /// Moves the object the specified number of characters to the right.
            /// </summary>
            /// <param name="Extend">
            /// Optional. Determines whether the moved text is collapsed or not. The default
            /// is false.
            /// </param>
            /// <param name="Count">
            /// Optional. Represents the number of characters to move to the right. The default
            /// is 1.
            /// </param>
        }

        /// <summary>Clears any unnamed bookmarks in the current text buffer line.</summary>        
        this.ClearBookmark = function () { }

        /// <summary>Collapses the selected text to the active po.</summary>        
        this.Collapse = function () { }

        /// <summary>Copies the selected text to the clipboard.</summary>        
        this.Copy = function () { }

        /// <summary>Copies the selected text to the clipboard and deletes it from its original location.</summary>        
        this.Cut = function () { }

        /// <summary>Deletes the selected text.</summary>
        /// Parameters:
        ///   Count:
        /// Optional. Represents the number of characters to delete.
        this.Delete = function (Count) { }

        /// <summary>Deletes a specified number of characters to the left of the active po.</summary>
        /// Parameters:
        ///   Count:
        /// Optional. Represents the number of characters to delete.
        this.DeleteLeft = function (Count) { }

        /// <summary>Deletes the empty characters (white space) horizontally or vertically around  the current location in the text buffer.</summary>    

        /// Parameters:
        ///   Direction:
        /// Optional. A EnvDTE.vsWhitespaceOptions constant that determines how and where
        /// to remove empty spaces.
        this.DeleteWhitespace = function (Direction) { }

        /// <summary>Inserts text, overwriting the existing text.</summary>
        /// Parameters:
        ///   Text:
        /// Required. Represents the text to insert.
        this.DestructiveInsert = function (Text) { }

        /// <summary>Moves the object to the end of the document.</summary>
        /// Parameters:
        ///   Extend:
        /// Optional. Determines whether the moved text is collapsed or not. The default
        /// is false.
        this.EndOfDocument = function (Extend) { }

        /// <summary>Moves the object to the end of the current line.</summary>
        /// Parameters:
        ///   Extend:
        /// Optional. Determines whether the moved text is collapsed or not. The default
        /// is false.
        this.EndOfLine = function (Extend) { }

        /// <summary>Searches for the given pattern from the active poto the end of the document.</summary>
        /// Parameters:
        ///   Pattern:
        /// Required. The text to find.

        ///   vsFindOptionsValue:
        /// One of the EnvDTE.vsFindOptions values.

        ///   Tags:
        /// Optional. If the matched pattern is a regular expression containing tagged subexpressions,
        /// then the Tags argument contains a collection of EnvDTE.TextRange objects, one
        /// for each tagged subexpression.

        /// Returns:
        /// A Boolean value indicating true if the pattern is found, false if not.
        this.FindPattern = function (Pattern, vsFindOptionsValue, Tags) { }

        /// <summary>Searches for the given text from the active poto the end of the document.</summary>
        /// Parameters:
        ///   Pattern:
        /// Required. The text to find.

        ///   vsFindOptionsValue:
        /// Optional. A EnvDTE.vsFindOptions constant indicating the search options to use.

        /// Returns:
        /// A Boolean value indicating true if the text is found, false if not.
        this.FindText = function (Pattern, vsFindOptionsValue) { }

        /// <summary>Moves to the beginning of the indicated line and selects the line if requested.</summary>
        /// Parameters:
        ///   Line:
        /// Required. The line number to go to, beginning at one.

        ///   Select:
        /// Optional. Indicates whether the target line should be selected. The default is
        /// false.
        this.GotoLine = function (Line, Select) { }

        /// <summary>Indents the selected lines by the given number of indentation levels.</summary>
        /// Parameters:
        ///   Count:
        /// Optional. The number of display indent levels to indent each line in the selected
        /// text. The default is 1.
        this.Indent = function (Count) { }

        /// <summary>Inserts the given at the current insertion point.</summary>
        /// Parameters:
        ///   Text:
        /// The text to insert.

        ///   vsInsertFlagsCollapseToEndValue:
        /// One of the EnvDTE.vsInsertFlags values indicating how to insert the text.
        this.Insert = function (Text, vsInsertFlagsCollapseToEndValue) { }

        /// <summary>Inserts the contents of the specified file at the current location in the buffer.</summary>
        /// Parameters:
        ///   File:
        /// Required. The name of the file to insert into the text buffer.
        this.InsertFromFile = function (File) { }

        /// <summary>Moves the insertion poof the text selection down the specified number of lines.</summary>     

        /// Parameters:
        ///   Extend:
        /// Optional. Determines whether the line in which the insertion pois moved is
        /// highlighted. The default is false.

        ///   Count:
        /// Optional. Indicates how many lines down to move the insertion po. The default
        /// value is 1.
        this.LineDown = function (Extend, Count) { }

        /// <summary>Moves the insertion poof the text selection up the specified number of lines.</summary>
        /// Parameters:
        ///   Extend:
        /// Optional. Determines whether the line in which the insertion pois moved is
        /// highlighted. The default is false.

        ///   Count:
        /// Optional. Indicates how many lines up to move the insertion po. The default
        /// is 1.
        this.LineUp = function (Extend, Count) { }

        /// <summary>Microsoft Internal Use Only.</summary>
        /// Parameters:
        ///   Line:
        /// The line number.

        ///   Column:
        /// The column number.

        ///   Extend:
        /// true if the move is extended, otherwise false.
        this.MoveTo = function (Line, Column, Extend) { }

        /// <summary>Moves the active poto the given 1-based absolute character offset.</summary>
        /// Parameters:
        ///   Offset:
        /// Required. A character index from the start of the document, starting at one

        ///   Extend:
        /// Optional. Default. A Boolean value to extend the current selection. If
        /// Extend is true, then the active end of the selection moves to the location while
        /// the anchor end remains where it is. Otherwise, both ends are moved to the specified
        /// location. This argument applies only to the EnvDTE.TextSelection object.
        this.MoveToAbsoluteOffset = function (Offset, Extend) { }

        /// <summary>Moves the active poto the indicated display column.</summary>
        /// Parameters:
        ///   Line:
        /// Required. A EnvDTE.vsGoToLineOptions constant representing the line offset, starting
        /// at one, from the beginning of the buffer.

        ///   Column:
        /// Required. Represents the virtual display column, starting at one, that is the
        /// new column location.

        ///   Extend:
        /// Optional. Determines whether the moved text is collapsed or not. The default
        /// is false.
        this.MoveToDisplayColumn = function (Line, Column, Extend) { }

        /// <summary>Moves the active poto the given position.</summary>
        /// Parameters:
        ///   Line:
        /// Required. The line number to move to, beginning at one. Line may also be one
        /// of the constants from EnvDTE.vsGoToLineOptions.

        ///   Offset:
        /// Required. The character index position in the line, starting at one.

        ///   Extend:
        /// Optional. Default. A Boolean value to extend the current selection. If
        /// Extend is true, then the active end of the selection moves to the location, while
        /// the anchor end remains where it is. Otherwise, both ends are moved to the specified
        /// location. This argument applies only to the EnvDTE.TextSelection object.
        this.MoveToLineAndOffset = function (Line, Offset, Extend) { }

        /// <summary>Moves the active poto the given position.</summary>
        /// Parameters:
        ///   Po:
        /// Required. The location in which to move the character.

        ///   Extend:
        /// Optional. Default. Determines whether to extend the current selection.
        /// If Extend is true, then the active end of the selection moves to the location,
        /// while the anchor end remains where it is. Otherwise, both ends are moved to the
        /// specified location. This argument applies only to the EnvDTE.TextSelection object.
        this.MoveToPoint = function (TextPoPo, Extend) { }

        /// <summary>Inserts a line break character at the active po.</summary>
        /// Parameters:
        ///   Count:
        /// Optional. Represents the number of NewLine characters to insert.
        this.NewLine = function (Count) { }

        /// <summary>Moves to the location of the next bookmark in the document.</summary>
        /// Returns:
        /// A Boolean value indicating true if the insertion pomoves to the next bookmark,
        /// false if otherwise.
        this.NextBookmark = function () { }

        /// <summary>Creates an outlining section based on the current selection.</summary>        
        this.OutlineSection = function () { }

        /// <summary>Fills the current line in the buffer with empty characters (white space) to the given column.</summary> 

        /// Parameters:
        ///   Column:
        /// Required. The number of columns to pad, starting at one.
        this.PadToColumn = function (Column) { }

        /// <summary>Moves the active poa specified number of pages down in the document, scrolling the view.</summary>    

        /// Parameters:
        ///   Extend:
        /// Optional. Determines whether the moved text is collapsed or not. The default
        /// is false.

        ///   Count:
        /// Optional. Represents the number of pages to move down. The default value is 1.
        this.PageDown = function (Extend, Count) { }

        /// <summary>Moves the active poa specified number of pages up in the document, scrolling the view.</summary>    

        /// Parameters:
        ///   Extend:
        /// Optional. Determines whether the moved text is collapsed or not. The default
        /// is false.

        ///   Count:
        /// Optional. Represents the number of pages to move up. The default value is 1.
        this.PageUp = function (Extend, Count) { }

        /// <summary>Inserts the clipboard contents at the current location.</summary>        
        this.Paste = function () { }

        /// <summary>Moves the text selection to the location of the previous bookmark in the document.</summary>
        /// Returns:
        /// A Boolean true if the text selection moves to a previous bookmark, false if not.
        this.PreviousBookmark = function () { }

        /// <summary>Replaces matching text throughout an entire text document.</summary>
        /// Parameters:
        ///   Pattern:
        /// Required. The to find.

        ///   Replace:
        /// Required. The text to replace each occurrence of Pattern.

        ///   vsFindOptionsValue:
        /// Optional. A EnvDTE.vsFindOptions constant indicating the behavior of EnvDTE.TextSelection.ReplacePattern(System.String,System.String,System.Int32,EnvDTE.TextRanges@),
        /// such as how to search, where to begin the search, whether to search forward or
        /// backward, and the case sensitivity.

        ///   Tags:
        /// Optional. A EnvDTE.TextRanges collection. If the matched text pattern is a regular
        /// expression and contains tagged subexpressions, then Tags contains a collection
        /// of EnvDTE.EditPoobjects, one for each tagged subexpression.

        /// Returns:
        /// A Boolean value.
        this.ReplacePattern = function (Pattern, Replace, vsFindOptionsValue, Tags) { }

        /// <summary>Microsoft Internal Use Only.</summary>
        /// Parameters:
        ///   Pattern:
        /// The pattern to find.

        ///   Replace:
        /// The with which to replace the found text.

        ///   vsFindOptionsValue:
        /// The find flags.

        /// Returns:
        /// true if the text was replaced, otherwise false.
        this.ReplaceText = function (Pattern, Replace, vsFindOptionsValue) { }

        /// <summary>Selects the entire document.</summary>        
        this.SelectAll = function () { }

        /// <summary>Selects the line containing the active po.</summary>        
        this.SelectLine = function () { }

        /// <summary>Sets an unnamed bookmark on the current line in the buffer.</summary>        
        this.SetBookmark = function () { }

        /// <summary>Formats the selected lines of text based on the current language.</summary>        
        this.SmartFormat = function () { }

        /// <summary>Moves the insertion poto the beginning of the document.</summary>
        /// Parameters:
        ///   Extend:
        /// Optional. Determines whether the text between the current location of the insertion
        /// poand the beginning of the document is highlighted or not. The default value
        /// is false.
        this.StartOfDocument = function (Extend) { }

        /// <summary>Moves the object to the beginning of the current line.</summary>
        /// Parameters:
        ///   Where:
        /// Optional. A EnvDTE.vsStartOfLineOptions constant representing where the line
        /// starts.
        ///   Extend:
        /// Optional. Determines whether the moved text is collapsed or not. The default
        /// is false.
        this.StartOfLine = function (Where, Extend) { }

        /// <summary>Exchanges the position of the active and the anchor pos.</summary>        
        this.SwapAnchor = function () { }

        /// <summary>Converts spaces to tabs in the selection according to your tab settings.</summary>        
        this.Tabify = function () { }

        /// <summary>Removes indents from the selected text by the number of indentation levels given.</summary>
        /// Parameters:
        ///   Count:
        /// Optional. The number of display indent levels to remove from each line in the
        /// selected text. The default is 1.
        this.Unindent = function (Count) { }

        /// <summary>Converts tabs to spaces in the selection according to the user's tab settings.</summary>        
        this.Untabify = function () { }

        /// <summary>Moves the selected text left the specified number of words.</summary>
        /// Parameters:
        ///   Extend:
        /// Optional. Determines whether the moved text is collapsed or not. The default
        /// is false.
        ///   Count:
        /// Optional. Represents the number of words to move left. The default value is 1.
        this.WordLeft = function (Extend, Count) { }

        /// <summary>Moves the selected text right the specified number of words.</summary>
        /// Parameters:
        ///   Extend:
        /// Optional. Determines whether the moved text is collapsed or not. The default
        /// is false.
        ///   Count:
        /// Optional. Represents the number of words to move right. The default value is
        /// 1.
        this.WordRight = function (Extend, Count) { }
    }
}
var _dte = new function() {

    /// Represents a document in the environment open for editing.
    this.Document = new function () {
        /// Gets the collection containing the EnvDTE.Document object.
        this.Collection = "";

        /// Gets the top-level extensibility object.
        this.DTE = "";

        /// Gets the Extender category ID (CATID) for the object.
        this.ExtenderCATID = "";

        /// Gets a list of available Extenders for the object.
        this.ExtenderNames = "";

        /// Gets the full path and name of the object's file.
        this.FullName = "";

        /// Microsoft Internal Use Only.
        this.IndentSize = "";

        /// Gets a GUID string indicating the kind or type of the object.
        this.Kind = "";

        /// Microsoft Internal Use Only.
        this.Language = "";

        /// Gets the name of the EnvDTE.Document.
        this.Name = "";

        /// Gets the path, without file name, for the directory containing the document.
        this.Path = "";

        /// Gets the EnvDTE.ProjectItem object associated with the EnvDTE.Document object.
        this.ProjectItem = "";

        /// Microsoft Internal Use Only.
        this.ReadOnly = "";

        /// Returns true if the object has not been modified since last being saved or opened.
        this.Saved = "";

        /// Gets an object representing the current selection on the EnvDTE.Document.
        this.Selection = new function () {
            return _dte2.Selection;
        }

        /// Microsoft Internal Use Only.
        this.TabSize = "";

        /// Microsoft Internal Use Only.
        this.Type = "";

        /// Gets a EnvDTE.Windows collection containing the windows that display in the object.
        this.Windows = "";

        /// Moves the focus to the current item.
        this.Activate = function () { }

        /// Microsoft Internal Use Only.
        this.ClearBookmarks = function () { }

        this.Close = function () {
            /// <summary>Closes the open document and optionally saves it, or closes and destroys the window.</summary>
            /// <params>Optional. A vsSaveChanges constant that determines whether to save an item or items (vsSaveChanges.vsSaveChangesNo, vsSaveChanges.vsSaveChangesPrompt, vsSaveChanges.vsSaveChangesYes.</params>
        }

        /// Closes the open document and optionally saves it, or closes and destroys the
        this.NewWindow = function () { }

        /// Returns an interface or object that can be accessed at run time by name.
        this.PrintOut = function () { }

        /// Re-executes the last action that was undone by the EnvDTE.Document.Undo method
        this.Redo = function () { }

        /// Saves the document.
        this.Save = function () { }

        /// Microsoft Internal Use Only.
        this.Undo = function () {
        }
    }

    /// Contains all Document objects in the environment, each representing an open document.
    this.Documents = new function () {
        /// Gets a value indicating the number of objects in the EnvDTE.Documents collection.
        this.Count = "";

        /// Gets the top-level extensibility object.
        this.DTE = "";

        /// Gets the immediate parent object of a EnvDTE.Documents collection.
        this.Parent = "";

        /// <summary>Closes all open documents in the environment and optionally saves them.</summary>
        /// <param name="Save">Optional. A EnvDTE.vsSaveChanges constant representing how to react to changes made to documents</param>
        this.CloseAll = function (Save) { }

        /// Returns an enumerator for items in the collection.
        this.GetEnumerator = function () { }

        this.Item = function (index) {
            /// <summary>Returns an indexed member of a EnvDTE.Documents collection.</summary>
            /// <param name="index">Required. The index of the item to return.</param>

            return _dte.Document;
        }

        /// Saves all documents currently open in the environment.
        this.SaveAll = function () { }

    }

    this.ActiveWindow = new function () { }

    this.AddIns = new function () { }

    this.CommandBars = new function () { }
}

/// The top-level object in the Visual Studio automation object model.
var dte = new function () {
    /// Gets the active document.
    this.ActiveDocument = new function () {
        return _dte.Document;
    }

    /// Gets an array of currently selected projects.
    this.ActiveSolutionProjects = "";

    /// Returns the currently active window, or the top-most window if no others are active.
    this.ActiveWindow = new function () {
        return _dte.ActiveWindow;
    }

    /// Gets the EnvDTE.AddIns collection, which contains all currently available Add-ins.
    this.AddIns = new function () {
        return _dte.AddIns;
    }

    /// Gets a reference to the development environment's command bars.
    this.CommandBars = new function () {
        return _dte.CommandBars;
    }

    /// Gets a string representing the command line arguments.
    this.CommandLineArguments = new function () {
    }

    /// Returns the EnvDTE.Commands collection.
    this.Commands = new function () {
    }

    /// Gets a EnvDTE.ContextAttributes collection which allows automation clients to add new attributes to the current selected items in the Dynamic Help window and provide contextual help for the additional attributes.
    this.ContextAttributes = new function () {
    }

    /// Gets the debugger objects.
    this.Debugger = new function () {
    }

    /// Gets the display mode, either MDI or Tabbed Documents.
    this.DisplayMode = new function () {
    }

    /// Gets the collection of open documents in the development environment.
    this.Documents = new function () {
        return _dte.Documents;
    }

    /// Gets the top-level extensibility object.
    this.DTE = new function () {
    }

    /// Gets a description of the edition of the environment.
    this.Edition = new function () {
    }

    /// Gets a reference to the EnvDTE.Events object.
    this.Events = new function () {
    }

    /// Microsoft Internal Use Only.
    this.FileName = new function () {
    }

    /// Gets the EnvDTE.Find object that represents global text find operations.
    this.Find = new function () {
    }

    /// Gets the full path and name of the object's file.
    this.FullName = new function () {
    }

    /// Gets the EnvDTE.Globals object that contains Add-in values that may be saved in the solution (.sln) file, the project file, or in the user's profile data.
    this.Globals = new function () {
    }

    /// Gets the EnvDTE.ItemOperations object.
    this.ItemOperations = new function () {
    }

    /// Gets the ID of the locale in which the development environment is running.
    this.LocaleID = new function () {
    }

    /// Gets a EnvDTE.Window object representing the main development environment window.
    this.MainWindow = new function () {
    }

    /// Gets the mode of the development environment, either debug or design.
    this.Mode = new function () {
    }

    /// Sets or gets the name of the EnvDTE._DTE object.
    this.Name = new function () {
    }

    /// Gets the EnvDTE.ObjectExtenders object.
    this.ObjectExtenders = new function () {
    }

    /// Gets a string with the path to the root of the Visual Studio registry settings.
    this.RegistryRoot = new function () {
    }

    /// Gets a collection containing the items currently selected in the environment.
    this.SelectedItems = new function () {
    }

    /// Gets the EnvDTE.Solution object that represents all open projects in the current instance of the environment and allows access to the build objects.
    this.Solution = new function () {
    }

    /// Gets a EnvDTE.SourceControl object that allows you to manipulate the source code control state of the file behind the object.
    this.SourceControl = new function () {
    }

    /// Gets the EnvDTE.StatusBar object, representing the status bar on the main development environment window.
    this.StatusBar = new function () {
    }

    /// Gets or sets whether UI should be displayed during the execution of automation code.
    this.SuppressUI = new function () {
    }

    /// Gets the global EnvDTE.UndoContext object.
    this.UndoContext = new function () {
    }

    /// Sets or gets a value indicating whether the environment was launched by a user or by automation.
    this.UserControl = new function () {
    }

    /// Gets the host application's version number.
    this.Version = new function () {
    }

    /// Gets the EnvDTE.WindowConfigurations collection, representing all available window configurations.
    this.WindowConfigurations = new function () {
    }

    /// Gets a EnvDTE.Windows collection containing the windows that display in the object.
    this.Windows = new function () {
    }

    this.ExecuteCommand = function (CommandName, CommandArgs) {
        /// <summary>Executes the specified command.</summary>
        /// <param name="CommandName" type="String">Required. The name of the command to invoke.</param>
        /// <param name="CommandArgs" type="String">Optional. A string containing the same arguments you would supply if you were invoking the command from the Command window. If a string is supplied, it is passed to the command line as the command's first argument and is parsed to form the various arguments for the command. This is similar to how commands are invoked in the Command window.</param>
    }
}

/// The top-level object in the Macros object model
var Macro = new function () {
    this.InsertText = function (text) {
        /// <summary>Inserts the given at the current insertion point.</summary>
        /// <param name="text" type="String">Required. The text to insert</param>
    };
};