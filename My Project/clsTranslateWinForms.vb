' IDEMS International
' Copyright (C) 2021
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License 
' along with this program.  If not, see <http://www.gnu.org/licenses/>.
Imports System.ComponentModel
Imports System.Data.SQLite
Imports System.Reflection
Imports System.Windows.Forms

'''------------------------------------------------------------------------------------------------
''' <summary>   
''' Provides utility functions to translate the text in WinForm objects (e.g. menu items, forms and 
''' controls) to a different natural language (e.g. to French). 
''' <para>
''' This class uses an SQLite database to translate text items to a new language. The database must 
''' contain the following tables:
''' <code>
''' CREATE TABLE "form_controls" (
'''	"form_name"	TEXT,
'''	"control_name"	TEXT,
'''	"id_text"	TEXT NOT NULL,
'''	PRIMARY KEY("form_name", "control_name")
''' )
''' </code><code>
''' CREATE TABLE "translations" (
'''	"id_text"	TEXT,
'''	"language_code"	TEXT,
'''	"translation"	TEXT NOT NULL,
'''	PRIMARY KEY("id_text", "language_code")
''' )
''' </code></para><para>
''' For example, if the 'form_controls' table contains a row with the values 
''' {'frmMain', 'mnuFile', 'File'}, 
''' then the 'translations' table should have a row for each supported language, e.g. 
''' {'File', 'en', 'File'}, {'File', 'fr', 'Fichier'}.
''' </para><para>
''' Note: This class is intended to be used solely as a 'static' class (i.e. contains only shared 
''' members, cannot be instantiated and cannot be inherited from).
''' In order to enforce this (and prevent developers from using this class in an unintended way), 
''' the class is declared as 'NotInheritable` and the constructor is declared as 'Private'.</para>
''' </summary>
'''------------------------------------------------------------------------------------------------
Public NotInheritable Class clsTranslateWinForms

    '''--------------------------------------------------------------------------------------------
    ''' <summary> 
    ''' Declare constructor 'Private' to prevent instantiation of this class (see class comments 
    ''' for more details). 
    ''' </summary>
    '''--------------------------------------------------------------------------------------------
    Private Sub New()
    End Sub

    '''--------------------------------------------------------------------------------------------
    ''' <summary>
    '''     Translates all the text in form <paramref name="clsForm"/> into language 
    '''     <paramref name="strLanguageCode"/> using the translations in database 
    '''     <paramref name="strDataSource"/>.
    '''     All the form's (sub)controls and (sub) menus are translated.     
    ''' </summary>
    '''
    ''' <param name="clsForm">          The WinForm form to translate. </param>
    ''' <param name="strDataSource">    The path of the SQLite '.db' file that contains the
    '''                                 translation database. </param>
    ''' <param name="strLanguageCode">      The language code to translate to (e.g. 'fr' for French). 
    '''                                 </param>
    '''
    ''' <returns>   If an exception is thrown, then returns the exception text; else returns 
    '''             'Nothing'. </returns>
    '''--------------------------------------------------------------------------------------------
    Public Shared Function TranslateForm(clsForm As Form, strDataSource As String,
                                         strLanguageCode As String) As String
        If IsNothing(clsForm) OrElse String.IsNullOrEmpty(strDataSource) OrElse
                String.IsNullOrEmpty(strLanguageCode) Then
            Return ("Developer Error: Illegal parameter passed to TranslateForm (language: " &
                   strLanguageCode & ", source: " & strDataSource & ").")
        End If

        Dim dctComponents As Dictionary(Of String, Component) = New Dictionary(Of String, Component)
        clsWinFormsComponents.FillDctComponentsFromControl(clsForm, dctComponents)
        Return TranslateDctComponents(dctComponents, clsForm.Name, strDataSource, strLanguageCode)

    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   
    '''    Translates all the (sub)menu items in <paramref name="clsMenuItems"/> into language
    '''    <paramref name="strLanguageCode"/> using the translations in database
    '''    <paramref name="strDataSource"/>.
    ''' </summary>
    '''
    ''' <param name="strParentName">    The menu's parent control. </param>
    ''' <param name="clsMenuItems">     The (sub)menu items to translate. </param>
    ''' <param name="strDataSource">    The path of the SQLite '.db' file that contains the
    '''                                 translation database. </param>
    ''' <param name="strLanguageCode">      The language code to translate to (e.g. 'fr' for French).
    '''                                 </param>
    '''
    ''' <returns>   If an exception is thrown, then returns the exception text; else returns 
    '''             'Nothing'. </returns>
    '''--------------------------------------------------------------------------------------------
    Public Shared Function TranslateMenuItems(strParentName As String, clsMenuItems As ToolStripItemCollection,
                                              strDataSource As String, strLanguageCode As String) As String
        If IsNothing(clsMenuItems) OrElse String.IsNullOrEmpty(strParentName) OrElse
                String.IsNullOrEmpty(strDataSource) OrElse String.IsNullOrEmpty(strLanguageCode) Then
            Return ("Developer Error: Illegal parameter passed to TranslateMenuItems (language: " &
                   strLanguageCode & ", source: " & strDataSource & ", parent: " & strParentName & " ).")
        End If

        Dim dctComponents As Dictionary(Of String, Component) = New Dictionary(Of String, Component)
        clsWinFormsComponents.FillDctComponentsFromMenuItems(clsMenuItems, dctComponents)

        Return TranslateDctComponents(dctComponents, strParentName, strDataSource, strLanguageCode)
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   
    '''    Returns <paramref name="strText"/> translated into <paramref name="strLanguageCode"/>. 
    '''    <para>
    '''    Translations can be bi-directional (e.g. from English to French or from French to English).
    '''    If <paramref name="strText"/> is already in the current language, or if no translation 
    '''    can be found, then returns <paramref name="strText"/>.         
    '''    </para></summary>
    '''
    ''' <param name="strText">          The text to translate. </param>
    ''' <param name="strDataSource">    The path of the SQLite '.db' file that contains the
    '''                                 translation database. </param>
    ''' <param name="strLanguageCode">      The language code to translate to (e.g. 'fr' for French).
    '''                                 </param>
    '''
    ''' <returns>   <paramref name="strText"/> translated into <paramref name="strLanguageCode"/>. </returns>
    '''--------------------------------------------------------------------------------------------
    Public Shared Function GetTranslation(strText As String, strDataSource As String,
                                          strLanguageCode As String) As String
        Dim strTranslation As String = ""
        Try
            'connect to the SQLite database that contains the translations
            Dim clsBuilder As New SQLiteConnectionStringBuilder With {
                .FailIfMissing = True,
                .DataSource = strDataSource}
            Using clsConnection As New SQLiteConnection(clsBuilder.ConnectionString)
                clsConnection.Open()
                strTranslation = GetDynamicTranslation(strText, strLanguageCode, clsConnection)
                clsConnection.Close()
            End Using
        Catch e As Exception
            Return e.Message & Environment.NewLine &
                    "A problem occured attempting to translate string '" & strText &
                    "' to language " & strLanguageCode & " using database " & strDataSource & "."
        End Try
        Return strTranslation
    End Function


    '''--------------------------------------------------------------------------------------------
    ''' <summary>
    '''     Attempts to translate all the text in <paramref name="dctComponents"/>
    '''     to <paramref name="strLanguageCode"/>.
    '''     Opens database <paramref name="strDataSource"/> and reads in all translations for the 
    '''     <paramref name="strControlName"/> control for target language <paramref name="strLanguageCode"/>.
    '''     For each translation in the database, attempts to find the corresponding component in 
    '''     <paramref name="dctComponents"/>. If found, then it translates the text to the target 
    '''     language. If a component has a tool tip, then it also translates the tool tip.
    ''' </summary>
    '''
    ''' <param name="dctComponents">    [in,out] The dictionary of translatable components. </param>
    ''' <param name="strControlName">   The name of the form or menu used to populate the dictionary. </param>
    ''' <param name="strDataSource">    The path of the SQLite '.db' file that contains the
    '''                                 translation database. </param>
    ''' <param name="strLanguageCode">      The language code to translate to (e.g. 'fr' for French). </param>
    '''
    ''' <returns>
    '''     If an exception is thrown, then returns the exception text; else returns 'Nothing'.
    ''' </returns>
    '''--------------------------------------------------------------------------------------------
    Private Shared Function TranslateDctComponents(ByRef dctComponents As Dictionary(Of String, Component),
                                                   strControlName As String,
                                                   strDataSource As String,
                                                   strLanguageCode As String) As String
        'Create a list of all the tool tip objects associated with this (sub)dialog
        'Note: Normally, a (sub)dialog wil only have a single tool tip object. This stores the 
        '      tool tips for all the components in the (sub)dialog.
        Dim lstToolTips As New List(Of ToolTip)
        For Each clsDctEntry As KeyValuePair(Of String, Component) In dctComponents
            'TODO for efficiency, we assume that only forms and user controls have tool tip objects.
            'Also allow other component types to have tool tip objects?
            If Not (TypeOf clsDctEntry.Value Is Form) AndAlso Not (TypeOf clsDctEntry.Value Is UserControl) Then
                Continue For
            End If
            'Tool tip objects are stored in the form's private list of components. There is no 
            '    public access. Therefore we need multiple steps to get the list of tool tip objects.
            Dim clsType As Type = clsDctEntry.Value.GetType()
            Dim clsFieldInfo As FieldInfo = clsType?.GetField("components", BindingFlags.Instance Or BindingFlags.NonPublic)
            Dim clsContainer As Container = clsFieldInfo?.GetValue(clsDctEntry.Value)
            Dim lstToolTipsCtrl As List(Of ToolTip) = clsContainer?.Components.OfType(Of ToolTip).ToList()
            If lstToolTipsCtrl IsNot Nothing Then
                For Each clsToolTip As ToolTip In lstToolTipsCtrl
                    lstToolTips.Add(clsToolTip)
                Next
            End If
        Next

        Try
            'connect to the SQLite database that contains the translations
            Dim clsBuilder As New SQLiteConnectionStringBuilder With {
                .FailIfMissing = True,
                .DataSource = strDataSource}
            Using clsConnection As New SQLiteConnection(clsBuilder.ConnectionString)
                clsConnection.Open()
                Using clsCommand As New SQLiteCommand(clsConnection)

                    'get all static translations for the specified form and language
                    clsCommand.CommandText =
                            "SELECT control_name, form_controls.id_text, translation " &
                            "FROM form_controls, translations WHERE form_name = '" & strControlName &
                            "' AND language_code = '" & strLanguageCode &
                            "' AND form_controls.id_text = translations.id_text"
                    Dim clsReader As SQLiteDataReader = clsCommand.ExecuteReader()
                    Using clsReader

                        'for each translation row
                        While clsReader.Read()

                            'find the component in the dictionary
                            Dim strComponentName As String = clsReader.GetString(0)
                            Dim strIdText As String = clsReader.GetString(1)
                            Dim strTranslation As String = clsReader.GetString(2)
                            Dim clsComponent As Component = clsWinFormsComponents.GetComponent(dctComponents, strComponentName)

                            'if component not found then continue to next translation row
                            If clsComponent Is Nothing Then
                                Continue While
                            End If

                            'translate the component's text to the new language
                            If TypeOf clsComponent Is Control Then
                                Dim clsControl As Control = DirectCast(clsComponent, Control)
                                clsControl.Text = strTranslation
                            ElseIf TypeOf clsComponent Is ToolStripItem Then
                                Dim clsMenuItem As ToolStripItem = DirectCast(clsComponent, ToolStripItem)
                                clsMenuItem.Text = strTranslation
                            Else
                                MsgBox("Developer Error: Translation dictionary entry (" & strComponentName & ") contained unexpected value type.")
                                Exit While
                            End If
                            TranslateToolTip(lstToolTips, clsComponent, strLanguageCode, clsConnection)
                        End While

                    End Using
                    Using clsReader
                        'get all controls with dynamic translations for the specified form
                        clsCommand.CommandText = "SELECT control_name FROM form_controls WHERE form_name = '" & strControlName &
                                                 "' AND id_text = 'ReplaceWithDynamicTranslation'"
                        clsReader = clsCommand.ExecuteReader()

                        'for each control with a dynamic translation
                        While (clsReader.Read())

                            'translate the component's text to the new language
                            Dim strComponentName As String = clsReader.GetString(0)
                            Dim clsComponent As Component = Nothing
                            If dctComponents.TryGetValue(strComponentName, clsComponent) Then
                                If TypeOf clsComponent Is Control Then 'currently we only dynamically translate controls
                                    Dim clsControl As Control = DirectCast(clsComponent, Control)
                                    clsControl.Text = GetDynamicTranslation(clsControl.Text, strLanguageCode, clsConnection)
                                End If
                                TranslateToolTip(lstToolTips, clsComponent, strLanguageCode, clsConnection)
                            End If

                        End While
                    End Using
                End Using
                clsConnection.Close()
            End Using
        Catch e As Exception
            Return e.Message & Environment.NewLine &
                    "A problem occured attempting to translate to language " & strLanguageCode &
                    " using database " & strDataSource & "."
        End Try
        Return Nothing
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   
    '''   
    '''  If <paramref name="clsComponent"/> has a tool tip, then converts the tool tip into 
    '''  <paramref name="strLanguage"/>.
    '''  For controls, it searches in <paramref name="lstToolTips"/> for the control's tool tip. 
    '''  Tool bar buttons are not controls and their tool tips are not stored in 
    '''  <paramref name="lstToolTips"/>. Therefore, for tool bar buttons, if the tool bar button 
    '''  has tool tip text, then it translates the tool tip text directly.
    ''' </summary>
    '''
    ''' <param name="lstToolTips">      The tool tip object(s) that ay contain <paramref name="clsComponent"/>'s tool tip text. </param>
    ''' <param name="clsComponent">     The component that may have tool tip text. </param>
    ''' <param name="strLanguage">      The language code to translate to (e.g. 'fr' for French). </param>
    ''' <param name="clsConnection">    An open connection to the SQLite translation database. </param>
    '''--------------------------------------------------------------------------------------------
    Private Shared Sub TranslateToolTip(lstToolTips As List(Of ToolTip), clsComponent As Component, strLanguage As String, clsConnection As SQLiteConnection)

        If TypeOf clsComponent Is Control Then
            For Each clsToolTip As ToolTip In lstToolTips
                Dim clsControl As Control = DirectCast(clsComponent, Control)
                Dim strToolTip As String = clsToolTip.GetToolTip(clsControl)
                If Not String.IsNullOrEmpty(strToolTip) Then
                    clsToolTip.SetToolTip(clsControl, GetDynamicTranslation(strToolTip, strLanguage, clsConnection))
                    Exit For
                End If
            Next
        ElseIf TypeOf clsComponent Is ToolStripItem Then 'else if component is a tool bar button
            'Tool bar buttons are not controls and their tool tips are not stored in the form's tool tip object
            '    So we need to translate their tool tip text directly 
            Dim clsMenuItem As ToolStripItem = DirectCast(clsComponent, ToolStripItem)
            Dim strToolTip As String = clsMenuItem.ToolTipText
            If Not String.IsNullOrEmpty(strToolTip) Then
                clsMenuItem.ToolTipText = GetDynamicTranslation(strToolTip, strLanguage, clsConnection)
            End If
        End If

    End Sub

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   
    '''    Returns <paramref name="strText"/> translated into <paramref name="strLanguage"/>. 
    '''    <para>
    '''    Translations can be bi-directional (e.g. from English to French or from French to English).
    '''    If <paramref name="strText"/> is already in the current language, or if no translation 
    '''    can be found, then returns <paramref name="strText"/>.         
    '''    </para></summary>
    '''
    ''' <param name="strText">          The text to translate. </param>
    ''' <param name="strLanguage">      The language code to translate to (e.g. 'fr' for French).
    '''                                 </param>
    ''' <param name="clsConnection">    An open connection to the SQLite translation database. </param>
    '''
    ''' <returns>   <paramref name="strText"/> translated into <paramref name="strLanguage"/>. </returns>
    '''--------------------------------------------------------------------------------------------
    Private Shared Function GetDynamicTranslation(strText As String, strLanguage As String, clsConnection As SQLiteConnection) As String
        If String.IsNullOrEmpty(strText) Then
            Return ""
        End If

        Using clsCommand As New SQLiteCommand(clsConnection)

            'in the translation text, convert any single quotes to make them suitable for the SQL command
            strText = strText.Replace("'", "''")

            'get all translations for the specified form and language
            'Note: The second `SELECT` is needed because we may sometimes need to translate  
            '      translated text back to the original text (e.g. from French to English when 
            '      the dialog language toggle button is clicked).
            clsCommand.CommandText = "SELECT translation FROM translations WHERE language_code = '" &
                                     strLanguage & "' AND id_text = '" & strText & "' OR (language_code = '" &
                                     strLanguage & "' AND id_text = " &
                                     "(SELECT id_text FROM translations WHERE translation = '" & strText & "'))"
            Dim clsReader As SQLiteDataReader = clsCommand.ExecuteReader()
            Using clsReader
                'return the translation text
                If clsReader.Read() Then
                    Return clsReader.GetString(0)
                End If
            End Using
        End Using
        'if no translation text was found then return original text unchanged
        Return strText
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Returns translations from database <paramref name="strDataSource"/>, for language 
    ''' <paramref name="strLanguageCode"/> and ID of text to translate <paramref name="strIdText"/>.
    ''' If <paramref name="strLanguageCode"/> is not specified then returns translations for all 
    ''' languages.
    ''' If <paramref name="strIdText"/> is not specified, then returns all translations for the 
    ''' specified language(s).
    ''' The returned data table has 3 columns: <code>id_text, language_code, translation</code>
    ''' </summary>
    ''' <param name="strDataSource">The database file path </param>
    ''' <param name="strLanguageCode">Optional. Only returns translations for this language. If not 
    '''                               specified then returns translations for all languages</param>
    ''' <param name="strIdText">Optional. The ID of the text to translate. If not specified, then 
    '''                         returns all translations for the specified language(s).</param>
    ''' <returns>Translations from database <paramref name="strDataSource"/>, for language 
    '''      <paramref name="strLanguageCode"/> and text to translate <paramref name="strIdText"/>.
    '''      </returns>
    '''--------------------------------------------------------------------------------------------
    Public Shared Function GetTranslations(strDataSource As String, Optional strLanguageCode As String = "", Optional strIdText As String = "") As DataTable
        Dim clsDataTableTranslations As New DataTable
        Try
            'connect to the SQLite database that contains the translations
            Dim clsBuilder As New SQLiteConnectionStringBuilder With {
                .FailIfMissing = True,
                .DataSource = strDataSource}
            Using clsConnection As New SQLiteConnection(clsBuilder.ConnectionString)
                clsConnection.Open()

                Dim clsDataAdapter As New SQLiteDataAdapter
                Dim strSqlWhereClause As String = ""
                Using cmdSelect As New SQLiteCommand(clsConnection)

                    If Not String.IsNullOrEmpty(strLanguageCode) Then
                        strSqlWhereClause = " WHERE language_code = @language_code"
                        cmdSelect.Parameters.Add(New SQLiteParameter("language_code", strLanguageCode))
                    End If

                    If Not String.IsNullOrEmpty(strIdText) Then
                        strSqlWhereClause = If(strSqlWhereClause = "", " WHERE ", " AND ")
                        strSqlWhereClause &= "id_text = @id_text"
                        cmdSelect.Parameters.Add(New SQLiteParameter("id_text", strIdText))
                    End If

                    cmdSelect.CommandText =
                        "SELECT id_text, language_code, translation FROM translations " & strSqlWhereClause
                    clsDataAdapter.SelectCommand = cmdSelect
                    clsDataAdapter.Fill(clsDataTableTranslations)

                End Using

                clsConnection.Close()
            End Using
        Catch e As Exception
            Throw New Exception("Error. Could NOT get translations.", e)
        End Try
        Return clsDataTableTranslations
    End Function


End Class
