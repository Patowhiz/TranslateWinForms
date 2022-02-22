
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
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports System.ComponentModel
Imports System.Data.SQLite
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq


'''------------------------------------------------------------------------------------------------
''' <summary>   
''' Provides utility functions to help create a translations database that can be used by other 
''' classes in this package.
''' For example: 
''' <list type="bullet">
'''     <item><description>
'''             Update the 'translations' database table with the texts from a CrowdIn format JSON 
'''             file.
'''     </description></item><item><description>
'''             Recursively traverse a component hierarchy and returns the parent, name and 
'''             associated text of each component.
'''     </description></item><item><description>
'''             Populate the 'form_controls' and 'translations' database table.
'''     </description></item><item><description>
'''             Update the 'form_controls' database table based on the specifications in the 
'''             'translateIgnore.txt' file.
'''     </description></item><item><description>
'''             Save the text to be translated as a CrowdIn format JSON file.
'''     </description></item>
''' </list>
''' <para>
''' The database must 
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
Public NotInheritable Class clsGenerateTranslations

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
    ''' Updates the 'form_controls' table in the <paramref name="strDataSource"/> database based on 
    ''' the rows in <paramref name="clsDatatableControls"/>.
    ''' For each row in <paramref name="clsDatatableControls"/>, the function will delete the 
    ''' corresponding table row (if it exists) and then insert a new, updated row.
    ''' This function also updates the `TranslateWinForm` database based on the specifications in 
    ''' the <paramref name="strTranslateIgnoreFilePath"/> file.
    ''' </summary>
    ''' <param name="strDataSource">The file path to the sqlite database file</param>
    ''' <param name="clsDatatableControls">The form controls datatable with 3 columns; 
    ''' <code>form_name, control_name, id_text.</code></param>
    ''' <param name="strTranslateIgnoreFilePath">Optional. The file path to the translate ignore file. 
    ''' If passed translate ignore controls will also be processed.</param>
    ''' <returns>The number of successful updates.</returns>
    '''--------------------------------------------------------------------------------------------
    Public Shared Function UpdateFormControlsTable(strDataSource As String, clsDatatableControls As DataTable, Optional strTranslateIgnoreFilePath As String = "") As Integer
        Dim iRowsUpdated As Integer = 0
        Try
            'connect to the database and execute the SQL command 
            Dim clsBuilder As New SQLiteConnectionStringBuilder With {
                    .FailIfMissing = True,
                    .DataSource = strDataSource}
            Using clsConnection As New SQLiteConnection(clsBuilder.ConnectionString)
                clsConnection.Open()
                'todo. do batch execution for optimal performance
                For Each row As DataRow In clsDatatableControls.Rows

                    Dim paramFormName As New SQLiteParameter("form_name", row.Field(Of String)(0))
                    Dim paramControlName As New SQLiteParameter("control_name", row.Field(Of String)(1))
                    Dim paramIdText As New SQLiteParameter("id_text", row.Field(Of String)(2))

                    'delete record if exists first 
                    Dim sqlDeleteCommand As String = "DELETE FROM form_controls WHERE form_name = @form_name AND control_name=@control_name"
                    Using cmdDelete As New SQLiteCommand(sqlDeleteCommand, clsConnection)
                        cmdDelete.Parameters.Add(paramFormName)
                        cmdDelete.Parameters.Add(paramControlName)
                        cmdDelete.ExecuteNonQuery()
                    End Using

                    'insert the new record
                    Dim sqlInsertCommand As String = "INSERT INTO form_controls (form_name,control_name,id_text) VALUES (@form_name,@control_name,@id_text)"
                    Using cmdInsert As New SQLiteCommand(sqlInsertCommand, clsConnection)
                        cmdInsert.Parameters.Add(paramFormName)
                        cmdInsert.Parameters.Add(paramControlName)
                        cmdInsert.Parameters.Add(paramIdText)
                        iRowsUpdated += cmdInsert.ExecuteNonQuery()
                    End Using
                Next

                clsConnection.Close()
            End Using

        Catch e As Exception
            Throw New Exception("Error: Could NOT update the form_controls database table", e)
        End Try

        If iRowsUpdated <> clsDatatableControls.Rows.Count Then
            Throw New Exception("Error: Could NOT save all form controls to the form_controls table. Rows saved: " & iRowsUpdated)
        End If

        If Not String.IsNullOrEmpty(strTranslateIgnoreFilePath) Then
            SetFormControlsToTranslateIgnore(strDataSource, strTranslateIgnoreFilePath)
        End If

        Return iRowsUpdated
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Updates the 'translations' database table with the texts in the passed datatable
    ''' </summary>   
    ''' <param name="strDataSource">The database file path</param>
    ''' <param name="clsDatatableTranslations">The translations datatable with 3 columns; 
    ''' <code>id_text, language_code, translation.</code></param>
    ''' <returns>The number of successful updates.</returns>
    '''--------------------------------------------------------------------------------------------
    Public Shared Function UpdateTranslationsTable(strDataSource As String, clsDatatableTranslations As DataTable) As Integer
        Dim iRowsUpdated As Integer = 0
        Try
            Dim clsBuilder As New SQLiteConnectionStringBuilder With {
                    .FailIfMissing = True,
                    .DataSource = strDataSource}
            Using clsConnection As New SQLiteConnection(clsBuilder.ConnectionString)
                clsConnection.Open()

                'todo. this needs to be done as a batch execution for optimal performance
                For Each row As DataRow In clsDatatableTranslations.Rows
                    Dim paramIdText As New SQLiteParameter("id_text", row.Field(Of String)("id_text"))
                    Dim paramLangcode As New SQLiteParameter("language_code", row.Field(Of String)("language_code"))
                    Dim paramTranslation As New SQLiteParameter("translation", row.Field(Of String)("translation"))

                    'delete record if exists first 
                    Dim sqlDeleteCommand As String = "DELETE FROM translations WHERE id_text = @id_text AND language_code=@language_code"
                    Using cmdDelete As New SQLiteCommand(sqlDeleteCommand, clsConnection)
                        cmdDelete.Parameters.Add(paramIdText)
                        cmdDelete.Parameters.Add(paramLangcode)
                        cmdDelete.ExecuteNonQuery()
                    End Using

                    'insert the new record
                    Dim sqlInsertCommand As String = "INSERT INTO translations (id_text,language_code,translation) VALUES (@id_text,@language_code,@translation)"
                    Using cmdInsert As New SQLiteCommand(sqlInsertCommand, clsConnection)
                        cmdInsert.Parameters.Add(paramIdText)
                        cmdInsert.Parameters.Add(paramLangcode)
                        cmdInsert.Parameters.Add(paramTranslation)
                        iRowsUpdated += cmdInsert.ExecuteNonQuery()
                    End Using
                Next

                clsConnection.Close()
            End Using
        Catch e As Exception
            Throw New Exception("Error: Could NOT update the translations database table", e)
        End Try

        If iRowsUpdated <> clsDatatableTranslations.Rows.Count Then
            Throw New Exception("Error: Could NOT save all form id texts to the translations table. Rows saved: " & iRowsUpdated)
        End If

        Return iRowsUpdated
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Updates the 'translations' database table with the texts from controls in the passed datatable
    ''' </summary>   
    ''' <param name="strDataSource">The file path to the sqlite database</param>
    ''' <param name="clsDatatableControls">The form controls datatable with 3 columns; 
    ''' <code>form_name, control_name, id_text.</code></param>
    ''' <returns>The number of successful updates.</returns>
    '''--------------------------------------------------------------------------------------------
    Public Shared Function UpdateTranslationsTableFromControls(strDataSource As String, clsDatatableControls As DataTable) As Integer
        Return UpdateTranslationsTable(strDataSource, GetTranslationTextsFromControls(clsDatatableControls))
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Returns a datatable with all form controls from the list of forms passed
    ''' </summary>
    ''' <param name="lstForms">The forms to get the controls for translations</param>
    ''' <returns>A datatable with 3 columns; form_name, control_name, id_text. </returns>
    '''--------------------------------------------------------------------------------------------
    Public Shared Function GetControlsDatatable(lstForms As List(Of Form)) As DataTable
        Dim clsDatatableControls As New DataTable
        clsDatatableControls.Columns.Add("form_name", GetType(String))
        clsDatatableControls.Columns.Add("control_name", GetType(String))
        clsDatatableControls.Columns.Add("id_text", GetType(String))

        For Each frm As Form In lstForms
            Dim dctComponents As Dictionary(Of String, Component) = New Dictionary(Of String, Component)
            clsWinFormsComponents.FillDctComponentsFromControl(frm, dctComponents)

            For Each clsComponent In dctComponents
                Dim idText As String = ""
                If TypeOf clsComponent.Value Is Control Then
                    idText = GetActualTranslationText(DirectCast(clsComponent.Value, Control).Text)
                ElseIf TypeOf clsComponent.Value Is ToolStripItem Then
                    idText = GetActualTranslationText(DirectCast(clsComponent.Value, ToolStripItem).Text)
                End If
                'add row of form_name, control_name, id_text
                clsDatatableControls.Rows.Add(frm.Name, clsComponent.Key, idText)
            Next

            'Special case for radio buttons in panels: 
            '  Before the dialog is shown, each radio button is a direct child of the dialog 
            '  (e.g. 'dlg_Augment_rdoNewDataframe'). After the dialog is shown, the raio button becomes 
            '  a direct child of its parent panel.
            '  Therefore, we need to show the dialog before we traverse the dialog's control hierarchy.
            '  Unfortunately showing the dialog means that it has to be manually closed. So we only 
            '  show the dialog for this special case to save the developer from having to manually 
            '  close too many dialogs.
            '  TODO: launch each dialog in a new thread to avoid need for manual close?
            'If strTemp.ToLower().Contains("pnl") AndAlso strTemp.ToLower().Contains("rdo") Then
            '    'frmTemp.ShowDialog()
            '    frmTemp.Show()
            '    strTemp = GetControls(frmTemp)
            '    frmTemp.Close()
            'End If
        Next

        Return clsDatatableControls
    End Function

    'todo. can probably be improved futher to include "DoNotTranslate".
    'also the heuristics here should be defined at the product level. see issue #4
    Private Shared Function GetActualTranslationText(strText As String) As String
        If String.IsNullOrEmpty(strText) OrElse
            strText.Contains(vbCr) OrElse    'multiline text
            strText.Contains(vbLf) OrElse Not Regex.IsMatch(strText, "[a-zA-Z]") Then
            'Regex.IsMatch(strText, "CheckBox\d+$") OrElse 'CheckBox1, CheckBox2 etc. normally indicates dynamic translation
            'Regex.IsMatch(strText, "Label\d+$") OrElse 'Label1, Label2 etc. normally indicates dynamic translation

            'text that doesn't contain any letters (e.g. number strings)
            Return "ReplaceWithDynamicTranslation"
        End If
        Return strText
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the translation texts from a datatable that has the forms controls texts
    ''' </summary>
    ''' <param name="clsDatatableControls">The form controls datatable; form_name, control_name, id_text. </param>
    ''' <param name="strLangCode">Optional. The translations texts language code, default is 'en'</param>
    ''' <returns>The translations datatable with 3 columns; id_text, language_code, translation.</returns>
    '''--------------------------------------------------------------------------------------------
    Private Shared Function GetTranslationTextsFromControls(clsDatatableControls As DataTable, Optional strLangCode As String = "en") As DataTable
        'Fill translations table from the form controls table
        Dim clsDatatableTranslations As New DataTable
        ' Create 3 columns in the DataTable.
        clsDatatableTranslations.Columns.Add("id_text", GetType(String))
        clsDatatableTranslations.Columns.Add("language_code", GetType(String))
        clsDatatableTranslations.Columns.Add("translation", GetType(String))
        For Each row As DataRow In clsDatatableControls.Rows
            'ignore "ReplaceWithDynamicTranslation" id text
            If row.Field(Of String)("id_text") = "ReplaceWithDynamicTranslation" Then
                Continue For
            End If
            'add id_text, language_code, translation
            clsDatatableTranslations.Rows.Add(row.Field(Of String)("id_text"), strLangCode, row.Field(Of String)("translation"))
        Next
        Return clsDatatableTranslations
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   
    '''    Updates the `TranslateWinForm` database based on the specifications in the 
    '''    'translateIgnore.txt' file. This file provides a way to ignore specified WinForm 
    '''    controls when the application or dialog is translated into a different language.
    '''    <para>
    '''    For example, this file can be used to ensure that text that references pre-existing data 
    '''    or meta data (e.g. a file name, data frame name, column name, cell value etc.) stays the 
    '''    same, even when the rest of the dialog is translated into French or Portuguese.
    '''    </para><para>
    '''    This sub should be executed prior to each release to ensure that the `TranslateWinForm` 
    '''    database specifies all the controls to ignore during the translation.  </para> 
    ''' <param name="strDataSource">The database file path</param>
    ''' <param name="strTranslateIgnoreFilePath">The translate ignore file path</param>
    ''' <returns>The number of successful updates.</returns> 
    ''' </summary>
    '''--------------------------------------------------------------------------------------------
    Public Shared Function SetFormControlsToTranslateIgnore(strDataSource As String, strTranslateIgnoreFilePath As String) As Integer
        Dim iRowsUpdated As Integer = 0
        Dim lstIgnore As New List(Of String)
        Dim lstIgnoreNegations As New List(Of String)

        Try
            'For each line in the ignore file 
            Using clsReader As New StreamReader(strTranslateIgnoreFilePath)
                Do While clsReader.Peek() >= 0
                    Dim strIgnoreFileLine = clsReader.ReadLine().Trim()
                    If String.IsNullOrEmpty(strIgnoreFileLine) Then
                        Continue Do
                    End If

                    Select Case strIgnoreFileLine(0)
                        Case "#"
                        'Ignore comment lines
                        Case "!"
                            'Add negation pattern to negation list
                            lstIgnoreNegations.Add(strIgnoreFileLine.Substring(1)) 'remove leading '!'
                        Case Else
                            'Add pattern to ignore list
                            lstIgnore.Add(strIgnoreFileLine)
                    End Select
                Loop
            End Using
        Catch e As Exception
            Throw New Exception("Error: Could NOT process ignore file: " & strTranslateIgnoreFilePath, e)
        End Try
        'If the ignore file didn't contain any specifications, then it's probably an error
        'please note its not expected that the product developer will run this function
        'if no ignore specifications are defined in the file
        If lstIgnore.Count <= 0 AndAlso lstIgnoreNegations.Count <= 0 Then
            Throw New Exception("Error: The " & strTranslateIgnoreFilePath & " ignore file was processed. No ignore specifications were found. " &
                   "The database was not updated.")
            Return iRowsUpdated
        End If

        'create the SQL command to update the database
        Dim strSqlUpdate As String = "UPDATE form_controls SET id_text = 'DoNotTranslate' WHERE "

        If lstIgnore.Count > 0 Then
            strSqlUpdate &= "("
            For iListPos As Integer = 0 To lstIgnore.Count - 1
                strSqlUpdate &= If(iListPos > 0, " OR ", "")
                strSqlUpdate &= "control_name LIKE '" & lstIgnore.Item(iListPos) & "'"
            Next iListPos
            strSqlUpdate &= ")"
        End If

        If lstIgnoreNegations.Count > 0 Then
            strSqlUpdate &= If(lstIgnore.Count > 0, " AND ", "")
            strSqlUpdate &= "NOT ("
            For iListPos As Integer = 0 To lstIgnoreNegations.Count - 1
                strSqlUpdate &= If(iListPos > 0, " OR ", "")
                strSqlUpdate &= "control_name LIKE '" & lstIgnoreNegations.Item(iListPos) & "'"
            Next iListPos
            strSqlUpdate &= ")"
        End If

        Try
            'connect to the database and execute the SQL command
            Dim clsBuilder As New SQLiteConnectionStringBuilder With {
                    .FailIfMissing = True,
                    .DataSource = strDataSource}
            Using clsConnection As New SQLiteConnection(clsBuilder.ConnectionString)
                Using clsSqliteCmd As New SQLiteCommand(strSqlUpdate, clsConnection)
                    clsConnection.Open()
                    iRowsUpdated = clsSqliteCmd.ExecuteNonQuery()
                    clsConnection.Close()
                End Using
            End Using
        Catch e As Exception
            Throw New Exception("Error:Could NOT update translate ignore in the form_controls database table", e)
        End Try

        Return iRowsUpdated
    End Function


    '''--------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Updates the 'translations' database table with the texts from the CrowdIn JSON file
    ''' <para>Please note, example of the expected JSON format;
    ''' <code>
    ''' { "Cloud ht": "Cloud ht", "Temp": "Temp","Wrun": "Wrun", "Evap": "Évapo"}
    ''' </code></para>
    ''' </summary>
    ''' <param name="strDataSource">The database file path</param>
    ''' <param name="strJsonFilePath">The json file path</param>
    ''' <param name="strLanguageCode">The translations texts language code</param>
    ''' <returns>The number of successful updates.</returns>
    '''--------------------------------------------------------------------------------------------
    Public Shared Function UpdateTranslationsTableFromCrowdInJSONFile(strDataSource As String, strJsonFilePath As String, strLanguageCode As String) As Integer
        Dim clsDatatableTranslations As New DataTable
        'Create 3 columns in the DataTable.
        clsDatatableTranslations.Columns.Add("id_text", GetType(String))
        clsDatatableTranslations.Columns.Add("language_code", GetType(String))
        clsDatatableTranslations.Columns.Add("translation", GetType(String))

        Using reader As New StreamReader(strJsonFilePath)
            'read in the CrowdIn JSON object (CrowdIn file is a single big JSON object)
            Dim jsonObject As JObject = JToken.ReadFrom(New JsonTextReader(reader))
            'iterate through the json object properties and fill translations table 
            For Each jsonProperty As JProperty In jsonObject.Children(Of JProperty)
                clsDatatableTranslations.Rows.Add(jsonProperty.Name, strLanguageCode, jsonProperty.Value.ToString)
            Next
        End Using

        Return UpdateTranslationsTable(strDataSource, clsDatatableTranslations)
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Saves the translations as a json file.  
    ''' </summary>
    ''' <param name="strDataSource">The database file path</param>
    ''' <param name="strSaveFilePathName">The file path and name to save.</param>
    ''' <param name="strLanguageCode"></param>
    ''' <returns>Number of translations writen to the json file</returns>
    '''--------------------------------------------------------------------------------------------
    Public Shared Function WriteTranslationsToCrowdInJSONFile(strDataSource As String, strSaveFilePathName As String, Optional strLanguageCode As String = "") As Integer
        Dim clsDatatableTranslations As DataTable = clsTranslateWinForms.GetTranslations(strDataSource, strLanguageCode:=strLanguageCode)
        Dim jsonObject As New JObject

        'convert the translations into a crowdin json format
        For Each row As DataRow In clsDatatableTranslations.Rows
            jsonObject.Add(New JProperty(row.Field(Of String)("id_text"), row.Field(Of String)("translation")))
        Next

        'write the json object to a json file
        Using sw As New StreamWriter(strSaveFilePathName)
            sw.WriteLine(jsonObject.ToString())
            sw.Flush()
            sw.Close()
        End Using

        Return clsDatatableTranslations.Rows.Count
    End Function


    '''--------------------------------------------------------------------------------------------
    ''' <summary>   
    '''     Recursively traverses the <paramref name="clsControl"/> control hierarchy and returns a
    '''     string containing the parent, name and associated text of each control. The string is 
    '''     formatted as a comma-separated list suitable for importing into a database.
    ''' </summary>
    '''
    ''' <param name="clsControl">   The control to process (it's children and sub-children shall 
    '''                             also be processed recursively). </param>
    '''
    ''' <returns>   
    '''     A string containing the parent, name and associated text of each control in the 
    '''     hierarchy. The string is formatted as a comma-separated list suitable for importing 
    '''     into a database. </returns>
    '''--------------------------------------------------------------------------------------------
    Public Shared Function GetControlsAsCsv(clsControl As Control) As String
        If IsNothing(clsControl) Then
            Return ""
        End If

        Dim dctComponents As Dictionary(Of String, Component) = New Dictionary(Of String, Component)
        clsWinFormsComponents.FillDctComponentsFromControl(clsControl, dctComponents)

        Dim strControlsAsCsv As String = ""
        For Each clsComponent In dctComponents
            If TypeOf clsComponent.Value Is Control Then
                Dim clsTmpControl As Control = DirectCast(clsComponent.Value, Control)
                strControlsAsCsv &= clsControl.Name & "," & clsComponent.Key & "," & GetCsvText(clsTmpControl.Text) & vbCrLf
            ElseIf TypeOf clsComponent.Value Is ToolStripItem Then
                Dim clsMenuItem As ToolStripItem = DirectCast(clsComponent.Value, ToolStripItem)
                strControlsAsCsv &= clsControl.Name & "," & clsComponent.Key & "," & GetCsvText(clsMenuItem.Text) & vbCrLf
            Else
                Throw New Exception("Developer Error: Translation dictionary entry (" & clsControl.Name & "," & clsComponent.Key & ") contained unexpected value type.")
            End If
        Next

        Return strControlsAsCsv
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   
    '''     Recursively traverses the <paramref name="clsMenuItems"/> menu hierarchy and returns a 
    '''     string containing the parent, name and associated text of each (sub)menu option in 
    '''     <paramref name="clsMenuItems"/>. The string is formatted as a comma-separated list 
    '''     suitable for importing into a database.
    ''' </summary>
    '''
    ''' <param name="clsControl">        The WinForm control that is the parent of the menu. </param>
    ''' <param name="clsMenuItems">     The WinForm menu items to add to the return string. </param>
    '''
    ''' <returns>   
    '''     A string containing the parent and name of each (sub)menu option in
    '''     <paramref name="clsMenuItems"/>. The string is formatted as a comma-separated list
    '''     suitable for importing into a database. </returns>
    '''--------------------------------------------------------------------------------------------
    Public Shared Function GetMenuItemsAsCsv(clsControl As Control, clsMenuItems As ToolStripItemCollection) As String
        If IsNothing(clsControl) OrElse IsNothing(clsMenuItems) Then
            Return ""
        End If

        Dim dctComponents As Dictionary(Of String, Component) = New Dictionary(Of String, Component)
        clsWinFormsComponents.FillDctComponentsFromMenuItems(clsMenuItems, dctComponents)

        Dim strMenuItemsAsCsv As String = ""
        For Each clsComponent In dctComponents

            If TypeOf clsComponent.Value Is ToolStripItem Then
                Dim clsMenuItem As ToolStripItem = DirectCast(clsComponent.Value, ToolStripItem)
                strMenuItemsAsCsv &= clsControl.Name & "," & clsComponent.Key & "," & GetCsvText(clsMenuItem.Text) & vbCrLf
            Else
                Throw New Exception("Developer Error: Translation dictionary entry (" & clsControl.Name & "," & clsComponent.Key & ") contained unexpected value type.")
            End If

        Next
        Return strMenuItemsAsCsv
    End Function

    'todo. heuristics checked in this function need to be defined at the product level 
    '''--------------------------------------------------------------------------------------------
    ''' <summary>   
    '''    Decides whether <paramref name="strText"/> is likely to be changed during execution of 
    '''    the software. If no, then returns <paramref name="strText"/>. If yes, then returns 
    '''    'ReplaceWithDynamicTranslation'. It makes the decision based upon a set of heuristics.
    '''    <para>
    '''    This function is normally only used when creating a comma-separated list suitable for 
    '''    importing into a database. During program execution, the 'ReplaceWithDynamicTranslation'
    '''    text tells the library to dynamically try and translate the current text, rather than
    '''    looking up the static text associated with the control.</para></summary>
    '''
    ''' <param name="strText">  The text to assess. </param>
    '''
    ''' <returns>   Decides whether <paramref name="strText"/> is likely to be changed during 
    '''             execution of the software. If no, then returns <paramref name="strText"/>. 
    '''             If yes, then returns'ReplaceWithDynamicTranslation'. </returns>
    '''--------------------------------------------------------------------------------------------
    Private Shared Function GetCsvText(strText As String) As String
        If String.IsNullOrEmpty(strText) OrElse
                strText.Contains(vbCr) OrElse strText.Contains(vbLf) OrElse 'multiline text
                Regex.IsMatch(strText, "CheckBox\d+$") OrElse 'CheckBox1, CheckBox2 etc. normally indicates dynamic translation
                Regex.IsMatch(strText, "Label\d+$") OrElse 'Label1, Label2 etc. normally indicates dynamic translation
                Regex.IsMatch(strText, "ToolStrip\d+$") OrElse 'ToolStripSplitButton1, ToolStripSplitButton2 etc. normally indicates dynamic translation
                Not Regex.IsMatch(strText, "[a-zA-Z]") Then 'text that doesn't contain any letters (e.g. number strings)
            Return "ReplaceWithDynamicTranslation"
        End If
        Return strText
    End Function


End Class
