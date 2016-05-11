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
        /// <summary>Provides access to view-based editing operations and selected text.</summary>

        /// <field name="ActivePoint">Gets the current endpoint of the selection.<\field>
        this.ActivePoint = "";

        /// <field name="AnchorColumn">Microsoft Internal Use Only.<\field>
        this.AnchorColumn = "";

        /// <field name="AnchorPoint">Gets the origin point of the selection.<\field>
        this.AnchorPoint = "";

        /// <field name="BottomLine">Microsoft Internal Use Only.<\field>
        this.BottomLine = "";

        /// <field name="BottomPoint">Gets the point at the end of the selection.<\field>
        this.BottomPoint = "";

        /// <field name="CurrentColumn">Microsoft Internal Use Only.<\field>
        this.CurrentColumn = "";

        /// <field name="CurrentLine">Microsoft Internal Use Only.<\field>
        this.CurrentLine = "";

        /// <field name="DTE">Gets the top-level extensibility object.<\field>
        this.DTE = "";

        /// <field name="IsActiveEndGreater">Gets whether the active pois equal to the bottom po.<\field>
        this.IsActiveEndGreater = "";

        /// <field name="IsEmpty">Gets whether the anchor pois equal to the active po.<\field>
        this.IsEmpty = "";

        /// <field name="Mode">Sets or gets a value determining whether dragging the mouse selects in stream<\field>
        this.Mode = "";

        /// <field name="Parent">Gets the immediate parent object of a EnvDTE.TextSelection object.<\field>
        this.Parent = "";

        /// <field name="Text">Sets or gets the selected text.<\field>
        this.Text = "";

        /// <field name="TextPane">Gets the text pane that contains the selected text.<\field>
        this.TextPane = "";

        /// <field name="TextRanges">Gets a EnvDTE.TextRanges collection with one EnvDTE.TextRange object for each<\field>
        this.TextRanges = "";

        /// <field name="TopLine">Microsoft Internal Use Only.<\field>
        this.TopLine = "";

        /// <field name="TopPoint">Gets the top end of the selection.<\field>
        this.TopPoint = "";

        this.Backspace = function (Count) {
            /// <summary>Microsoft Internal Use Only.</summary>
            /// <param name="Count">The number of spaces.</param>
        }

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

        this.ClearBookmark = function () {
            /// <summary>Clears any unnamed bookmarks in the current text buffer line.</summary>        
        }

        this.Collapse = function () {
            /// <summary>Collapses the selected text to the active po.</summary>        
        }

        this.Copy = function () {
            /// <summary>Copies the selected text to the clipboard.</summary>        
        }

        this.Cut = function () {
            /// <summary>Copies the selected text to the clipboard and deletes it from its original location.</summary>        
        }

        this.Delete = function (Count) {
            /// <summary>Deletes the selected text.</summary>
            /// <param name="Count"> Optional. Represents the number of characters to delete.</param>
        }

        this.DeleteLeft = function (Count) {
            /// <summary>Deletes a specified number of characters to the left of the active po.</summary>
            /// <param name="Count"> Optional. Represents the number of characters to delete.</param>
        }

        this.DeleteWhitespace = function (Direction) {
            /// <summary>Deletes the empty characters (white space) horizontally or vertically around  the current location in the text buffer.</summary>
            /// <param name="Direction"> Optional. A EnvDTE.vsWhitespaceOptions constant that determines how and where to remove empty spaces.</param>
        }

        this.DestructiveInsert = function (Text) {
            /// <summary>Inserts text, overwriting the existing text.</summary>
            /// <param name="Text"> Required. Represents the text to insert.</param>
        }

        this.EndOfDocument = function (Extend) {
            /// <summary>Moves the object to the end of the document.</summary>
            /// <param name="Extend"> Optional. Determines whether the moved text is collapsed or not. The default is false.</param>
        }

        this.EndOfLine = function (Extend) {
            /// <summary>Moves the object to the end of the current line.</summary>
            /// <param name="Extend"> Optional. Determines whether the moved text is collapsed or not. The default is false.</param>
        }

        this.FindPattern = function (Pattern, vsFindOptionsValue, Tags) {
            /// <summary>Searches for the given pattern from the active point to the end of the document.</summary>
            /// <param name="Pattern"> Required. The text to find.</param>
            /// <param name="vsFindOptionsValue"> One of the EnvDTE.vsFindOptions values.</param>
            /// <param name="Tags"> Optional. If the matched pattern is a regular expression containing tagged subexpressions,
            /// then the Tags argument contains a collection of EnvDTE.TextRange objects, one
            /// for each tagged subexpression.
            /// </param>
            /// <returns>A Boolean value indicating true if the pattern is found, false if not.</returns>
        }

        this.FindText = function (Pattern, vsFindOptionsValue) {
            /// <summary>Searches for the given text from the active poto the end of the document.</summary>
            /// <param name="Pattern"> Required. The text to find.</param>
            /// <param name="vsFindOptionsValue"> Optional. A EnvDTE.vsFindOptions constant indicating the search options to use.</param>
            /// <returns>A Boolean value indicating true if the text is found, false if not.</returns>
        }

        this.GotoLine = function (Line, Select) {
            /// <summary>Moves to the beginning of the indicated line and selects the line if requested.</summary>
            /// <param name="Line"> Required. The line number to go to, beginning at one.</param>
            /// <param name="Select"> Optional. Indicates whether the target line should be selected. The default is
            /// false.</param>
        }

        this.Indent = function (Count) {
            /// <summary>Indents the selected lines by the given number of indentation levels.</summary>
            /// <param name="Count"> Optional. The number of display indent levels to indent each line in the selected
            /// text. The default is 1.</param>
        }

        this.Insert = function (Text, vsInsertFlagsCollapseToEndValue) {
            /// <summary>Inserts the given at the current insertion point.</summary>
            /// <param name="Text"> The text to insert.</param>
            /// <param name="vsInsertFlagsCollapseToEndValue"> One of the EnvDTE.vsInsertFlags values indicating how to insert the text.</param>
        }

        this.InsertFromFile = function (File) {
            /// <summary>Inserts the contents of the specified file at the current location in the buffer.</summary>
            /// <param name="File"> Required. The name of the file to insert into the text buffer.</param>
        }

        this.LineDown = function (Extend, Count) {
            /// <summary>Moves the insertion point of the text selection down the specified number of lines.</summary>     
            /// <param name="Extend"> Optional. Determines whether the line in which the insertion pois moved is
            /// highlighted. The default is false.</param>
            /// <param name="Count"> Optional. Indicates how many lines down to move the insertion po. The default
            /// value is 1.</param>
        }

        this.LineUp = function (Extend, Count) {
            /// <summary>Moves the insertion point of the text selection up the specified number of lines.</summary>
            /// <param name="Extend"> Optional. Determines whether the line in which the insertion pois moved is
            /// highlighted. The default is false.</param>
            /// <param name="Count"> Optional. Indicates how many lines up to move the insertion po. The default
            /// is 1.</param>
        }

        this.MoveTo = function (Line, Column, Extend) {
            /// <summary>Microsoft Internal Use Only.</summary>
            /// <param name="Line"> The line number.</param>
            /// <param name="Column"> The column number.</param>
            /// <param name="Extend"> true if the move is extended, otherwise false.</param>
        }

        this.MoveToAbsoluteOffset = function (Offset, Extend) {
            /// <summary>Moves the active poto the given 1-based absolute character offset.</summary>
            /// <param name="Offset"> Required. A character index from the start of the document, starting at one</param>
            /// <param name="Extend"> Optional. Default. A Boolean value to extend the current selection. If
            /// Extend is true, then the active end of the selection moves to the location while
            /// the anchor end remains where it is. Otherwise, both ends are moved to the specified
            /// location. This argument applies only to the EnvDTE.TextSelection object.
            /// </param>
        }

        this.MoveToDisplayColumn = function (Line, Column, Extend) {
            /// <summary>Moves the active poto the indicated display column.</summary>
            /// <param name="Line"> Required. A EnvDTE.vsGoToLineOptions constant representing the line offset, starting
            /// at one, from the beginning of the buffer.</param>
            /// <param name="Column"> Required. Represents the virtual display column, starting at one, that is the
            /// new column location.</param>
            /// <param name="Extend"> Optional. Determines whether the moved text is collapsed or not. The default
            /// is false.</param>
        }

        this.MoveToLineAndOffset = function (Line, Offset, Extend) {
            /// <summary>Moves the active poto the given position.</summary>
            /// <param name="Line"> Required. The line number to move to, beginning at one. Line may also be one
            /// of the constants from EnvDTE.vsGoToLineOptions.</param>
            /// <param name="Offset"> Required. The character index position in the line, starting at one.</param>
            /// <param name="Extend"> Optional. Default. A Boolean value to extend the current selection. If
            /// Extend is true, then the active end of the selection moves to the location, while
            /// the anchor end remains where it is. Otherwise, both ends are moved to the specified
            /// location. This argument applies only to the EnvDTE.TextSelection object.
            /// </param>
        }

        this.MoveToPoint = function (TextPoPo, Extend) {
            /// <summary>Moves the active poto the given position.</summary>
            /// <param name="Po"> Required. The location in which to move the character.</param>
            /// <param name="Extend"> Optional. Default. Determines whether to extend the current selection.
            /// If Extend is true, then the active end of the selection moves to the location,
            /// while the anchor end remains where it is. Otherwise, both ends are moved to the
            /// specified location. This argument applies only to the EnvDTE.TextSelection object.
            /// </param>
        }

        this.NewLine = function (Count) {
            /// <summary>Inserts a line break character at the active po.</summary>
            /// <param name="Count"> Optional. Represents the number of NewLine characters to insert.</param>
        }

        this.NextBookmark = function () {
            /// <summary>Moves to the location of the next bookmark in the document.</summary>
            /// <returns>A Boolean value indicating true if the insertion pomoves to the next bookmark,
            /// false if otherwise.</returns>
        }

        this.OutlineSection = function () {
            /// <summary>Creates an outlining section based on the current selection.</summary>        
        }

        this.PadToColumn = function (Column) {
            /// <summary>Fills the current line in the buffer with empty characters (white space) to the given column.</summary> 
            /// <param name="Column"> Required. The number of columns to pad, starting at one.</param>
        }

        this.PageDown = function (Extend, Count) {
            /// <summary>Moves the active poa specified number of pages down in the document, scrolling the view.</summary>    
            /// <param name="Extend"> Optional. Determines whether the moved text is collapsed or not. The default
            /// is false.</param>
            /// <param name="Count"> Optional. Represents the number of pages to move down. The default value is 1.</param>
        }

        this.PageUp = function (Extend, Count) {
            /// <summary>Moves the active poa specified number of pages up in the document, scrolling the view.</summary>    
            /// <param name="Extend"> Optional. Determines whether the moved text is collapsed or not. The default
            /// is false.</param>
            /// <param name="Count"> Optional. Represents the number of pages to move up. The default value is 1.</param>
        }

        this.Paste = function () {
            /// <summary>Inserts the clipboard contents at the current location.</summary>        
        }

        this.PreviousBookmark = function () {
            /// <summary>Moves the text selection to the location of the previous bookmark in the document.</summary>
            /// <returns>A Boolean true if the text selection moves to a previous bookmark, false if not.</returns>
        }

        this.ReplacePattern = function (Pattern, Replace, vsFindOptionsValue, Tags) {
            /// <summary>Replaces matching text throughout an entire text document.</summary>
            /// <param name="Pattern"> Required. The text to find.</param>
            /// <param name="Replace"> Required. The text to replace each occurrence of Pattern.</param>
            /// <param name="vsFindOptionsValue">
            /// Optional. A EnvDTE.vsFindOptions constant indicating the behavior of 
            /// EnvDTE.TextSelection.ReplacePattern(System.String,System.String,System.Int32,EnvDTE.TextRanges@),
            /// such as how to search, where to begin the search, whether to search forward or
            /// backward, and the case sensitivity.
            /// </param>
            /// <param name="Tags"> Optional. A EnvDTE.TextRanges collection. If the matched text pattern is a regular
            /// expression and contains tagged subexpressions, then Tags contains a collection
            /// of EnvDTE.EditPoobjects, one for each tagged subexpression.</param>
            /// <returns>A Boolean value.</returns>
        }

        this.ReplaceText = function (Pattern, Replace, vsFindOptionsValue) {
            /// <summary>Microsoft Internal Use Only.</summary>
            /// <param name="Pattern"> The pattern to find.</param>
            /// <param name="Replace"> The with which to replace the found text.</param>
            /// <param name="vsFindOptionsValue"> The find flags.</param>
            /// <returns>true if the text was replaced, otherwise false.</returns>
        }

        this.SelectAll = function () {
            /// <summary>Selects the entire document.</summary>        
        }

        this.SelectLine = function () {
            /// <summary>Selects the line containing the active po.</summary>        
        }

        this.SetBookmark = function () {
            /// <summary>Sets an unnamed bookmark on the current line in the buffer.</summary>        
        }

        this.SmartFormat = function () {
        }
        /// <summary>Formats the selected lines of text based on the current language.</summary>        

        this.StartOfDocument = function (Extend) {
            /// <summary>Moves the insertion poto the beginning of the document.</summary>
            /// <param name="Extend"> Optional. Determines whether the text between the current location of the insertion
            /// poand the beginning of the document is highlighted or not. The default value
            /// is false.</param>
        }

        this.StartOfLine = function (Where, Extend) {
            /// <summary>Moves the object to the beginning of the current line.</summary>
            /// <param name="Where"> Optional. A EnvDTE.vsStartOfLineOptions constant representing where the line
            /// starts.</param>
            /// <param name="Extend"> Optional. Determines whether the moved text is collapsed or not. The default
            /// is false.</param>
        }

        this.SwapAnchor = function () {
            /// <summary>Exchanges the position of the active and the anchor pos.</summary>        
        }

        this.Tabify = function () {
            /// <summary>Converts spaces to tabs in the selection according to your tab settings.</summary>        
        }

        this.Unindent = function (Count) {
            /// <summary>Removes indents from the selected text by the number of indentation levels given.</summary>
            /// <param name="Count"> Optional. The number of display indent levels to remove from each line in the
            /// selected text. The default is 1.</param>
        }

        this.Untabify = function () {
            /// <summary>Converts tabs to spaces in the selection according to the user's tab settings.</summary>        
        }

        this.WordLeft = function (Extend, Count) {
            /// <summary>Moves the selected text left the specified number of words.</summary>
            /// <param name="Extend"> Optional. Determines whether the moved text is collapsed or not. The default
            /// is false.</param>
            /// <param name="Count"> Optional. Represents the number of words to move left. The default value is 1.</param>
        }

        this.WordRight = function (Extend, Count) {
            /// <summary>Moves the selected text right the specified number of words.</summary>
            /// <param name="Extend"> Optional. Determines whether the moved text is collapsed or not. The default
            /// is false.</param>
            /// <param name="Count"> Optional. Represents the number of words to move right. The default value is
            /// 1.</param>
        }
    }
}

var _dte = new function () {

    /// Represents a document in the environment open for editing.
    this.Document = new function () {
        /// <field name="Collection">Gets the collection containing the EnvDTE.Document object.</field>
        this.Collection = "";

        /// <field name="DTE">Gets the top-level extensibility object.</field>
        this.DTE = "";

        /// <field name="ExtenderCATID">Gets the Extender category ID (CATID) for the object.</field>
        this.ExtenderCATID = "";

        /// <field name="ExtenderNames">Gets a list of available Extenders for the object.</field>
        this.ExtenderNames = "";

        /// <field name="FullName">Gets the full path and name of the object's file.</field>
        this.FullName = "";

        /// <field name="IndentSize">Microsoft Internal Use Only.</field>
        this.IndentSize = "";

        /// <field name="Kind">Gets a GUID string indicating the kind or type of the object.</field>
        this.Kind = "";

        /// <field name="Language">Microsoft Internal Use Only.</field>
        this.Language = "";

        /// <field name="Name">Gets the name of the EnvDTE.Document.</field>
        this.Name = "";

        /// <field name="Path">Gets the path, without file name, for the directory containing the document.</field>
        this.Path = "";

        /// <field name="ProjectItem">Gets the EnvDTE.ProjectItem object associated with the EnvDTE.Document object.</field>
        this.ProjectItem = "";

        /// <field name="ReadOnly">Microsoft Internal Use Only.</field>
        this.ReadOnly = "";

        /// <field name="Saved">Returns true if the object has not been modified since last being saved or opened.</field>
        this.Saved = "";

        /// <field name="Selection">Gets an object representing the current selection on the EnvDTE.Document.</field>
        this.Selection = new function () {
            return _dte2.Selection;
        }

        /// <field name="TabSize">Microsoft Internal Use Only.</field>
        this.TabSize = "";

        /// <field name="Type">Microsoft Internal Use Only.</field>
        this.Type = "";

        /// <field name="Windows">Gets a EnvDTE.Windows collection containing the windows that display in the object.</field>
        this.Windows = "";

        this.Activate = function () {
            /// <summary>Moves the focus to the current item.</summary>
        }

        this.ClearBookmarks = function () {
            /// <summary>Microsoft Internal Use Only.</summary>
        }

        this.Close = function (Save) {
            /// <summary>Closes the open document and optionally saves it, or closes and destroys the window.</summary>
            /// <param name="Save">Optional. A vsSaveChanges constant that determines whether to save an item or items (vsSaveChanges.vsSaveChangesNo, vsSaveChanges.vsSaveChangesPrompt, vsSaveChanges.vsSaveChangesYes.</params>
        }

        this.NewWindow = function () {
            /// <summary>Closes the open document and optionally saves it, or closes and destroys the</summary>
        }

        this.PrintOut = function () {
            /// <summary>Returns an interface or object that can be accessed at run time by name.</summary>
        }

        this.Redo = function () {
            /// <summary>Re-executes the last action that was undone by the EnvDTE.Document.Undo method</summary>
        }

        this.Save = function () {
            /// <summary>Saves the document.</summary>
        }

        this.Undo = function ()
            /// <summary>Microsoft Internal Use Only.</summary>
        {
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

        this.CloseAll = function (Save) {
            /// <summary>Closes all open documents in the environment and optionally saves them.</summary>
            /// <param name="Save">Optional. A EnvDTE.vsSaveChanges constant representing how to react to changes made to documents</param>
        }

        this.GetEnumerator = function () {
            /// Returns an enumerator for items in the collection.
        }

        this.Item = function (index) {
            /// <summary>Returns an indexed member of a EnvDTE.Documents collection.</summary>
            /// <param name="index">Required. The index of the item to return.</param>

            return _dte.Document;
        }

        this.SaveAll = function () {
            /// Saves all documents currently open in the environment.
        }
    }

    this.AddIns = new function () { }

    this.CommandBars = new function () { }

    this.Window = new function () {

    }

    this.Windows = new function () {
        /// Gets a value indicating the number of objects in the EnvDTE.Windows collection.
        this.Count = "";

        /// Gets the top-level extensibility object.
        this.DTE = "";

        /// Gets the immediate parent object of a EnvDTE.Windows collection.
        this.Parent = "";

        this.CreateLinkedWindowFrame = function (Window1, Window2, Link) {
            /// <summary>Creates a EnvDTE.Window object and places two windows in it.</summary>
            /// <param name="Window1">
            ///     Required. The first EnvDTE.Window object to link to the other.
            /// </param>
            /// <param name="Window2">
            ///     Required. The second EnvDTE.Window object to link to the other.
            /// </param>
            ///  <param name="Link"
            ///     Required. A EnvDTE.vsLinkedWindowType constant indicating the way the windows
            ///     should be joined.
            /// </param>
            /// <returns>
            ///     A EnvDTE.Window object.
            /// </returns>
        }

        this.CreateToolWindow = function (AddInInst, ProgID, Caption, GuidPosition, DocObj) {
            /// <summary>Creates a new tool window containing the specified EnvDTE.Document object or</summary>
            /// <param name="AddInInst">
            ///     Required. An EnvDTE.AddIn object whose lifetime determines the lifetime of the
            ///     tool window. </param>
            /// <param name="ProgID">
            ///     Required. The programmatic ID of the EnvDTE.Document object or ActiveX control.
            /// </param>
            /// <param name="Caption">
            ///     Required. The caption for the new tool window.
            /// </param>
            /// <param name="GuidPosition">
            ///     Required. A unique identifier for the new tool window, which can be used as an
            ///     index to EnvDTE.Windows.Item(System.Object).
            /// </param>
            /// <param name="DocObj">
            ///     Required. The EnvDTE.Document object or control to be hosted in the tool window.
            /// </param>
            /// <returns>
            ///     A EnvDTE.Window object.
            /// </returns>
        }

        this.GetEnumerator = function () {
            /// <summary>
            ///     Returns an enumeration for items in a collection.
            /// </summary>
            /// <returns>
            ///     An enumerator.
            /// </returns>
        }

        this.Item = function (index) {
            /// <summary>
            ///     Returns a EnvDTE.Window object in a EnvDTE.Windows collection.
            /// </summary>
            /// <param name="index">
            ///     Required. The index of the EnvDTE.Window object to return.
            /// </param>
            /// <returns>
            ///     A EnvDTE.Window object.
            /// </returns>

            return _dte.Window;
        }
    }
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
    this.FileName = "FileName";

    /// Gets the EnvDTE.Find object that represents global text find operations.
    this.Find = new function () {
    }

    /// Gets the full path and name of the object's file.
    this.FullName = "FullName";

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