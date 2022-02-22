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
Imports System.Windows.Forms
Imports System.Text.RegularExpressions


'''------------------------------------------------------------------------------------------------
''' <summary>
''' Provides utility functions for getting winforms components
''' <para>
''' Note: This class is intended to be used solely as a 'static' class (i.e. contains only shared 
''' members, cannot be instantiated and cannot be inherited from).
''' In order to enforce this (and prevent developers from using this class in an unintended way), 
''' the class is declared as 'NotInheritable` and the constructor is declared as 'Private'.</para>
''' </summary>
'''------------------------------------------------------------------------------------------------
Public NotInheritable Class ClsWinFormsComponents

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
    '''    Returns the component associated with <paramref name="strComponentName"/> in 
    '''    <paramref name="dctComponents"/>.
    '''    If an exact match is not found, then returns a component whose name is a superset of 
    '''    <paramref name="strComponentName"/>. 
    '''    If no match is found then returns Nothing.
    ''' </summary>
    '''
    ''' <param name="dctComponents">    The dictionary of translatable components. </param>
    ''' <param name="strComponentName"> Name of the component to search for. </param>
    '''
    ''' <returns>   The component associated with <paramref name="strComponentName"/> in
    '''    <paramref name="dctComponents"/>. If an exact match is not found, then returns a 
    '''    component whose name is a superset of <paramref name="strComponentName"/>.
    '''    If no match is found then returns Nothing. </returns>
    '''--------------------------------------------------------------------------------------------
    Public Shared Function GetComponent(dctComponents As Dictionary(Of String, Component), strComponentName As String) As Component
        Dim clsComponent As Component = Nothing

        If dctComponents.TryGetValue(strComponentName, clsComponent) OrElse
                Not Regex.IsMatch(strComponentName, "\w+_\w+_\w+") Then
            Return clsComponent
        End If

        'Edge Case: If the component name is not found in the dictionary, then look for a dictionary 
        '  key that is a superset of the component name. This check is needed because the 
        '  hierarchy of the WinForm controls may be slightly different at runtime, compared to 
        '  when the translation database was generated.
        '  For example, during testing we found cases for sub-dialog->tab->group->panel->radioButton 
        '  similar to:
        '    Database component name: sdgPlots_tbpPlotsOptions_tbpXAxis_ucrXAxis_grpAxisTitle_rdoSpecifyTitle
        '    Runtime  component name: sdgPlots_tbpPlotsOptions_tbpXAxis_ucrXAxis_grpAxisTitle_ucrPnlAxisTitle_rdoSpecifyTitle

        'split the full component name into 2 parts: parent names & child name
        '  e.g 'sdgPlots_tbpPlotsOptions_tbpXAxis_ucrXAxis_grpAxisTitle' & '_rdoSpecifyTitle'
        Dim iTmpIndex As Integer = strComponentName.LastIndexOf("_")
        Dim strParentNames As String = strComponentName.Substring(0, iTmpIndex)
        Dim strControlName As String = strComponentName.Substring(iTmpIndex)

        'if the dictionary contains a single key that matches <parents>_<other controls>_<child>,
        '  then return the component associated with that key
        Dim lstComponents As List(Of Component) =
            dctComponents.Where(Function(x) Regex.IsMatch(x.Key, strParentNames & "\w+_\w+" & strControlName)) _
            .Select(Function(x) x.Value).ToList()
        If lstComponents.Count = 1 Then
            clsComponent = lstComponents(0)
        End If

        Return clsComponent
    End Function


    '''--------------------------------------------------------------------------------------------
    ''' <summary>
    '''    Populates dictionary <paramref name="dctComponents"/> with the control 
    '''    <paramref name="clsControl"/> and its children.    
    '''    The dictionary can then be used to conveniently translate the control text (see other
    '''    functions and subs in this class).
    ''' </summary>
    '''
    ''' <param name="clsControl">       The control used to populate the dictionary. </param>
    ''' <param name="dctComponents">    [in,out] Dictionary to store the control and its children. 
    '''                                 </param>
    '''--------------------------------------------------------------------------------------------
    Public Shared Sub FillDctComponentsFromControl(clsControl As Control,
                                                   ByRef dctComponents As Dictionary(Of String, Component),
                                                   Optional strParentName As String = "")
        If IsNothing(clsControl) OrElse IsNothing(clsControl.Controls) OrElse IsNothing(dctComponents) Then
            Exit Sub
        End If

        'if control is valid, then add it to the dictionary
        Dim strControlName As String = ""
        If Not String.IsNullOrEmpty(clsControl.Name) Then
            strControlName = If(String.IsNullOrEmpty(strParentName), clsControl.Name, strParentName & "_" & clsControl.Name)
            If Not dctComponents.ContainsKey(strControlName) Then  'ignore components that are already in the dictionary
                dctComponents.Add(strControlName, clsControl)
            End If
        End If

        For Each ctlChild As Control In clsControl.Controls

            'Recursively process different types of menus and child controls
            If TypeOf ctlChild Is MenuStrip Then
                FillDctComponentsFromMenuItems(DirectCast(ctlChild, MenuStrip).Items, dctComponents)
            ElseIf TypeOf ctlChild Is ToolStrip Then
                FillDctComponentsFromMenuItems(DirectCast(ctlChild, ToolStrip).Items, dctComponents)
            ElseIf TypeOf ctlChild Is Control Then
                FillDctComponentsFromControl(ctlChild, dctComponents, strControlName)
            End If

        Next
    End Sub

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   
    '''    Populates dictionary <paramref name="dctComponents"/> with all the menu items, and 
    '''    sub-menu items in <paramref name="clsMenuItems"/>. 
    '''    The dictionary can then be used to conveniently translate the menu item text (see other 
    '''    functions and subs in this class).
    ''' </summary>
    '''
    ''' <param name="clsMenuItems">     The list of menu items to populate the dictionary. </param>
    ''' <param name="dctComponents">    [in,out] Dictionary to store the menu items. </param>
    '''--------------------------------------------------------------------------------------------
    Public Shared Sub FillDctComponentsFromMenuItems(clsMenuItems As ToolStripItemCollection, ByRef dctComponents As Dictionary(Of String, Component))
        If IsNothing(clsMenuItems) OrElse IsNothing(dctComponents) Then
            Exit Sub
        End If

        For Each clsMenuItem As ToolStripItem In clsMenuItems

            'if menu item is valid, then add it to the dictionary
            If Not String.IsNullOrEmpty(clsMenuItem.Name) AndAlso Not dctComponents.ContainsKey(clsMenuItem.Name) Then
                'ignore components that are already in the dictionary
                dctComponents.Add(clsMenuItem.Name, clsMenuItem)
            End If

            'Recursively process different types of sub-menu
            If TypeOf clsMenuItem Is ToolStripMenuItem Then
                Dim clsTmpMenuItem As ToolStripMenuItem = DirectCast(clsMenuItem, ToolStripMenuItem)
                If clsTmpMenuItem.HasDropDownItems Then
                    FillDctComponentsFromMenuItems(clsTmpMenuItem.DropDownItems, dctComponents)
                End If
            ElseIf TypeOf clsMenuItem Is ToolStripSplitButton Then
                Dim clsTmpMenuItem As ToolStripSplitButton = DirectCast(clsMenuItem, ToolStripSplitButton)
                If clsTmpMenuItem.HasDropDownItems Then
                    FillDctComponentsFromMenuItems(clsTmpMenuItem.DropDownItems, dctComponents)
                End If
            ElseIf TypeOf clsMenuItem Is ToolStripDropDownButton Then
                Dim clsTmpMenuItem As ToolStripDropDownButton = DirectCast(clsMenuItem, ToolStripDropDownButton)
                If clsTmpMenuItem.HasDropDownItems Then
                    FillDctComponentsFromMenuItems(clsTmpMenuItem.DropDownItems, dctComponents)
                End If
            End If

        Next
    End Sub


End Class
