Imports Autodesk.Navisworks.Api

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Recurse(ByVal oParentModItem As ModelItem, ByVal oParentNode As TreeNode)
        For Each oSubModItem As ModelItem In oParentModItem.Children
            Dim strNodeName As String = ""

            If String.IsNullOrEmpty(oSubModItem.DisplayName) Then
                If oSubModItem.IsInsert = True Then
                    strNodeName = oSubModItem.Children.ElementAt(0).DisplayName
                End If
                If oSubModItem.HasGeometry = True Then
                    strNodeName = oSubModItem.ClassDisplayName
                End If
            Else
                strNodeName = oSubModItem.DisplayName
            End If

            Dim oSubNode As TreeNode
            oSubNode = oParentNode.Nodes.Add(strNodeName)
            oSubNode.Tag = oSubModItem
            Recurse(oSubModItem, oSubNode)
        Next
    End Sub

    Private Sub TreeView1_NodeMouseClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles TreeView1.NodeMouseClick
        If e.Node.Tag <> vbNull Then
            Try
                Dim Item As ModelItem = e.Node.Tag
                Dim oCollection As ModelItemCollection
                Dim oDoc As Document = Application.ActiveDocument
                oCollection = New ModelItemCollection()
                oCollection.Add(Item)
                oDoc.CurrentSelection.CopyFrom(oCollection)

            Catch ex As Exception

            End Try
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim rootItem As ModelItem
        Me.Text = Application.ActiveDocument.CurrentFileName
        rootItem = Application.ActiveDocument.Models(0).RootItem
        Dim oRootNode As TreeNode
        oRootNode = TreeView1.Nodes.Add(rootItem.DisplayName)
        oRootNode.Tag = rootItem

        Recurse(rootItem, oRootNode)
    End Sub
End Class
