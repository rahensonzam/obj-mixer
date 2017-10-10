Imports System.IO

Public Class frmMain

    Dim models As New Dictionary(Of String, Model)
    Dim togglesSelectAll As New Dictionary(Of String, Boolean)
    Dim currentModel As String
    Dim outputModel As Model

    Public Structure loc3d
        Dim x As Decimal
        Dim y As Decimal
        Dim z As Decimal
    End Structure

    Public Structure loc2d
        Dim x As Decimal
        Dim y As Decimal
    End Structure

    Sub loadModel(ByVal filePath As String)
        Dim m As Model = readFile(filePath)
        If models.ContainsKey(m.Name) = False Then
            CheckedListBoxModels.Items.Add(m.Name)
            models.Add(m.Name, m)
            togglesSelectAll.Add(m.Name, False)
        Else
            MessageBox.Show("Model already added!", "Error", 0, MessageBoxIcon.Error)
        End If
    End Sub

    Function readFile(ByVal filePath As String) As Model

        Dim currentLine As String = ""
        ' Create Dictionaries to temporarily store model variables
        Dim tempVertices As New Dictionary(Of Integer, Vertex)
        Dim tempFaces As New Dictionary(Of Integer, Face)
        Dim tempGroups As New Dictionary(Of String, Group)

        Dim fileName As String = Path.GetFileNameWithoutExtension(filePath)

        ' Look for object files in the application directory
        Dim pathed As String = filePath

        'Read v vt vn
        If System.IO.File.Exists(pathed) = True Then
            Dim LoadReader As New System.IO.StreamReader(pathed)
            ' Vertex id counters
            Dim posId As Integer = 1
            Dim texId As Integer = 1
            Dim normId As Integer = 1
            Do While LoadReader.Peek() <> -1
                currentLine = LoadReader.ReadLine()
                If currentLine <> "" Then
                    If currentLine.Substring(0, 2) = "v " Then
                        addVertex(currentLine, posId, 0, tempVertices)
                        posId += 1
                    ElseIf currentLine.Substring(0, 3) = "vt " Then
                        addVertex(currentLine, texId, 1, tempVertices)
                        texId += 1
                    ElseIf currentLine.Substring(0, 3) = "vn " Then
                        addVertex(currentLine, normId, 2, tempVertices)
                        normId += 1
                    End If
                End If
            Loop
            LoadReader.Close()
        Else
            MessageBox.Show("File not found", "", 0, MessageBoxIcon.Exclamation)
            Return Nothing
        End If


        'Read faces and groups
        Dim faceId As Integer = 1
        Dim currentGroup As String = ""
        If System.IO.File.Exists(pathed) = True Then
            Dim LoadReader As New System.IO.StreamReader(pathed)
            Do While LoadReader.Peek() <> -1

                If LoadReader.Peek() = AscW("g") Then
                    Dim groupName As String = LoadReader.ReadLine().Substring(2)
                    If tempGroups.ContainsKey(groupName) = False Then
                        tempGroups.Add(groupName, New Group(groupName))
                    End If
                    currentGroup = groupName
                End If

                If LoadReader.Peek() = AscW("f") Then
                    currentLine = LoadReader.ReadLine()
                    addFace(currentLine, faceId, tempFaces)
                    tempGroups(currentGroup).Faces.Add(faceId)
                    faceId += 1
                Else
                    LoadReader.ReadLine()
                End If

            Loop
            LoadReader.Close()
        Else
            MessageBox.Show("File not found", "", 0, MessageBoxIcon.Exclamation)
            Return Nothing
        End If

        ' Create a new model with the filename as its name
        Return New Model(fileName, tempVertices, tempFaces, tempGroups)

    End Function

    Sub addVertex(ByVal vString As String, ByVal vId As Integer, ByVal vType As Integer, ByRef tempVertices As Dictionary(Of Integer, Vertex))
        Dim vArray() As String
        Dim tempLoc3d As loc3d
        Dim tempLoc2d As loc2d
        vArray = vString.Split(CChar(" "))
        If tempVertices.ContainsKey(vId) = False Then
            tempVertices.Add(vId, New Vertex(vId))
        End If

        'Dim vType As Integer

        'If vArray(0) = "v " Then
        '    vType = 0
        'ElseIf vArray(0) = "vt " Then
        '    vType = 1
        'ElseIf vArray(0) = "vn " Then
        '    vType = 2
        'End If

        Select Case vType
            Case 0
                tempLoc3d = tempVertices(vId).Pos
                tempLoc3d.x = CDec(vArray(1))
                tempLoc3d.y = CDec(vArray(2))
                tempLoc3d.z = CDec(vArray(3))
                tempVertices(vId).Pos = tempLoc3d
            Case 1
                tempLoc2d = tempVertices(vId).Tex
                tempLoc2d.x = CDec(vArray(1))
                tempLoc2d.y = CDec(vArray(2))
                tempVertices(vId).Tex = tempLoc2d
            Case 2
                tempLoc3d = tempVertices(vId).Norm
                tempLoc3d.x = CDec(vArray(1))
                tempLoc3d.y = CDec(vArray(2))
                tempLoc3d.z = CDec(vArray(3))
                tempVertices(vId).Norm = tempLoc3d
        End Select
    End Sub

    Sub addFace(ByVal fString As String, ByVal fId As Integer, ByRef tempFaces As Dictionary(Of Integer, Face))
        Dim fArray() As String
        Dim tempFaceVertices As New List(Of Integer())
        Dim tempFaceVertex() As String

        fArray = fString.Split(CChar(" "))

        For i As Integer = 1 To fArray.Length - 1
            tempFaceVertex = fArray(i).Split(CChar("/"))
            Dim tempFaceVertexC(tempFaceVertex.Length - 1) As Integer
            For j As Integer = 0 To tempFaceVertex.Length - 1
                tempFaceVertexC(j) = CInt(tempFaceVertex(j))
            Next j
            tempFaceVertices.Add(tempFaceVertexC)
        Next

        tempFaces.Add(fId, New Face(tempFaceVertices, fId))

    End Sub

    Function combineModels(ByVal selectedModels As List(Of String)) As Model
        statusText.Text = "Starting model combination..."
        Application.DoEvents()
        If selectedModels.Count = 1 Then
            ToolStripProgressBar1.Maximum = 3
            Return combine2Or1Models(models(selectedModels(0)), Nothing)
        ElseIf selectedModels.Count = 2 Then
            ToolStripProgressBar1.Maximum = 3
            Return combine2Or1Models(models(selectedModels(0)), models(selectedModels(1)))
        ElseIf selectedModels.Count > 2 Then
            Dim newModel As Model = models(selectedModels.Item(0))
            ToolStripProgressBar1.Maximum = selectedModels.Count * 3
            For i As Integer = 1 To selectedModels.Count - 1
                If i + 1 <= selectedModels.Count - 1 Then
                    Dim model As Model = models(selectedModels.Item(i + 1))
                    statusText.Text = "Combining with " & model.Name & "..."
                    Application.DoEvents()
                    newModel = combine2Or1Models(newModel, model)
                End If
            Next
            Return newModel
        Else
            Return Nothing
        End If
    End Function

    Function combine2Or1Models(ByVal m1 As Model, ByVal m2 As Model) As Model
        ToolStripProgressBar1.Value += 1
        Dim newModel As Model
        Dim tempVertices As New Dictionary(Of Integer, Vertex)
        Dim tempFaces As New Dictionary(Of Integer, Face)
        Dim tempGroups As New Dictionary(Of String, Group)

        ' Copy model 1 into the new model
        copyModel(m1, tempVertices, tempFaces, tempGroups)

        Dim lastFaceId As Integer
        For Each groupName As String In m1.SelectedGroups
            For Each faceId As Integer In m1.Groups(groupName).Faces
                If faceId > lastFaceId Then
                    lastFaceId = faceId
                End If
            Next
        Next

        'If model 1 only, skip model 2
        If m2 IsNot Nothing Then
            Dim lastVertexId As Integer = tempVertices.Count

            Dim tempVertices2 As New Dictionary(Of Integer, Vertex)
            Dim tempFaces2 As New Dictionary(Of Integer, Face)
            Dim tempGroups2 As New Dictionary(Of String, Group)

            ' Copy model 2 into a temp location
            copyModel(m2, tempVertices2, tempFaces2, tempGroups2)

            ' Shift all the vertices in model 2 down
            ' so they don't clash with vertices in model1
            shiftVIds(False, 0, -lastVertexId, tempVertices2, tempFaces2)
            ' Shift all the faces in model 2 down
            ' so they don't clash with faces in model1
            shiftFIds(-lastFaceId, tempFaces2, tempGroups2)

            ' Actually copy model2's vertices, faces and groups into the new model
            For Each pair As KeyValuePair(Of Integer, Vertex) In tempVertices2
                tempVertices.Add(pair.Key, pair.Value)
            Next
            For Each pair As KeyValuePair(Of Integer, Face) In tempFaces2
                tempFaces.Add(pair.Key, pair.Value)
            Next

            For Each pair As KeyValuePair(Of String, Group) In tempGroups2
                ' If this group already exists
                ' add my faces to it
                If tempGroups.ContainsKey(pair.Key) = False Then
                    tempGroups.Add(pair.Key, pair.Value)
                Else
                    Dim g As Group = tempGroups(pair.Key)
                    For Each faceId As Integer In pair.Value.Faces
                        g.Faces.Add(faceId)
                    Next
                    tempGroups(pair.Key) = g
                End If
            Next
        End If

        statusText.Text = "Compressing models..."
        Application.DoEvents()

        sortVerticesDictionary(tempVertices)
        compressModel(tempVertices, tempFaces)
        ToolStripProgressBar1.Value += 1
        Application.DoEvents()

        statusText.Text = "Removing duplicate/unused vertices..."
        Application.DoEvents()
        Dim numDuplicates As Integer = removeDuplicates(tempVertices, tempFaces)
        Dim numUnused As Integer = removeUnusedVertices(tempVertices, tempFaces)
        ToolStripProgressBar1.Value += 1
        Application.DoEvents()

        'Debug.Print("Removed " & numDuplicates & " duplicate vertices")
        'Debug.Print("Removed " & numUnused & " unused vertices")

        sortVerticesDictionary(tempVertices)
        compressModel(tempVertices, tempFaces)

        newModel = New Model("Result", tempVertices, tempFaces, tempGroups)
        Return newModel
    End Function

    Function loc3dEqual(ByVal l1 As loc3d, ByVal l2 As loc3d) As Boolean
        Dim str1 As String = CStr(l1.x) + CStr(l1.y) + CStr(l1.z)
        Dim str2 As String = CStr(l2.x) + CStr(l2.y) + CStr(l2.z)
        If str1.Equals(str2) = True Then
            Return True
        Else
            Return False
        End If
    End Function

    Function loc2dEqual(ByVal l1 As loc2d, ByVal l2 As loc2d) As Boolean
        Dim str1 As String = CStr(l1.x) + CStr(l1.y)
        Dim str2 As String = CStr(l2.x) + CStr(l2.y)
        If str1.Equals(str2) = True Then
            Return True
        Else
            Return False
        End If
    End Function

    Function findDuplicateVertices(ByVal v As Vertex, ByVal tempVertices As Dictionary(Of Integer, Vertex)) As Dictionary(Of Integer, Vertex)
        Dim duplicates As New Dictionary(Of Integer, Vertex)
        For Each pair As KeyValuePair(Of Integer, Vertex) In tempVertices
            Dim v2 As Vertex = pair.Value
            ' If this vertex is not the one we are finding duplicates of
            If v.Id <> v2.Id Then
                ' If the vertices are exactly equal
                If loc3dEqual(v.Pos, v2.Pos) = True And loc3dEqual(v.Norm, v2.Norm) = True And loc2dEqual(v.Tex, v2.Tex) Then
                    duplicates.Add(v2.Id, v2)
                End If
            End If
        Next
        Return duplicates
    End Function

    Function getFirstVId(ByVal vertices As Dictionary(Of Integer, Vertex)) As Integer
        Dim smallestVId As Integer = -1
        For Each pair As KeyValuePair(Of Integer, Vertex) In vertices
            Dim vertex As Vertex = pair.Value
            If vertex.Id < smallestVId Or smallestVId = -1 Then
                smallestVId = vertex.Id
            End If
        Next
        Return smallestVId
    End Function

    Function removeDuplicates(ByRef tempVertices As Dictionary(Of Integer, Vertex), ByRef tempFaces As Dictionary(Of Integer, Face)) As Integer
        Dim verticesToDelete As New List(Of Integer)
        For Each pair As KeyValuePair(Of Integer, Vertex) In tempVertices
            Dim vertex As Vertex = pair.Value
            Dim duplicates As Dictionary(Of Integer, Vertex) = findDuplicateVertices(vertex, tempVertices)
            ' If there is at least one duplicate
            If duplicates.Count > 0 Then
                ' Add this vertex when checking for first vertex id
                duplicates.Add(vertex.Id, vertex)
                Dim firstVId As Integer = getFirstVId(duplicates)
                If vertex.Id <> firstVId Then
                    vertex.NewId = firstVId
                    verticesToDelete.Add(vertex.Id)
                End If
            End If
        Next
        updateFaceVIds(tempVertices, tempFaces)
        For Each i As Integer In verticesToDelete
            tempVertices.Remove(i)
        Next
        Return verticesToDelete.Count
    End Function

    Function removeUnusedVertices(ByRef tempVertices As Dictionary(Of Integer, Vertex), ByRef tempFaces As Dictionary(Of Integer, Face)) As Integer
        Dim usedVertices As New List(Of Integer)
        For Each pair As KeyValuePair(Of Integer, Face) In tempFaces
            Dim face As Face = pair.Value
            For Each faceVertex As Integer() In face.FaceVertices
                For i As Integer = 0 To faceVertex.Length - 1
                    Dim vId As Integer = faceVertex(i)
                    If usedVertices.Contains(vId) = False Then
                        usedVertices.Add(vId)
                    End If
                Next i
            Next
        Next
        Dim verticesToDelete As New List(Of Integer)
        For Each pair As KeyValuePair(Of Integer, Vertex) In tempVertices
            Dim vertex As Vertex = pair.Value
            If usedVertices.Contains(vertex.Id) = False Then
                verticesToDelete.Add(vertex.Id)
            End If
        Next
        updateFaceVIds(tempVertices, tempFaces)
        For Each i As Integer In verticesToDelete
            tempVertices.Remove(i)
        Next
        Return verticesToDelete.Count
    End Function

    Sub copyModel(ByVal m As Model, ByRef tempVertices As Dictionary(Of Integer, Vertex), ByRef tempFaces As Dictionary(Of Integer, Face), ByRef tempGroups As Dictionary(Of String, Group))
        For Each groupName As String In m.SelectedGroups
            tempGroups.Add(groupName, New Group(groupName))
            Dim group As Group = m.Groups(groupName)
            tempGroups(groupName).Faces = group.Faces
            For Each faceId As Integer In group.Faces
                Dim face As Face = m.Faces(faceId)
                tempFaces.Add(faceId, face)
                For Each faceVertex As Integer() In face.FaceVertices
                    For i As Integer = 0 To faceVertex.Length - 1
                        Dim vId As Integer = faceVertex(i)
                        If tempVertices.ContainsKey(vId) = False Then
                            Dim vertex As Vertex = m.Vertices(vId)
                            tempVertices.Add(vertex.Id, vertex)
                        End If
                    Next i
                Next
            Next
        Next

        sortVerticesDictionary(tempVertices)
        compressModel(tempVertices, tempFaces)
    End Sub

    Sub sortVerticesDictionary(ByRef tempVertices As Dictionary(Of Integer, Vertex))
        Dim lastId As Integer = -1
        For Each pair As KeyValuePair(Of Integer, Vertex) In tempVertices
            If pair.Key > lastId Then
                lastId = pair.Key
            End If
        Next
        Dim newTempVertices As New Dictionary(Of Integer, Vertex)
        For i As Integer = 0 To lastId
            If tempVertices.ContainsKey(i) Then
                newTempVertices.Add(i, tempVertices(i))
            End If
        Next i
        tempVertices = newTempVertices
    End Sub

    Sub compressModel(ByRef tempVertices As Dictionary(Of Integer, Vertex), ByRef tempFaces As Dictionary(Of Integer, Face))
        Dim currentVertexId As Integer = 1
        Dim lastVertexId As Integer = 0
        ' Find the last id in temp vertices
        Dim greatestId As Integer = -1
        For Each pair As KeyValuePair(Of Integer, Vertex) In tempVertices
            If pair.Key > greatestId Then
                greatestId = pair.Key
            End If
        Next
        For i As Integer = 0 To greatestId
            If tempVertices.ContainsKey(i) = True Then
                currentVertexId = i
                Dim diff As Integer = currentVertexId - lastVertexId
                If diff > 1 Then
                    shiftVIds(True, lastVertexId, diff, tempVertices, tempFaces)
                    sortVerticesDictionary(tempVertices)
                    i = 0
                End If

                lastVertexId = currentVertexId
            End If
        Next i
    End Sub

    Sub updateFaceVIds(ByRef tempVertices As Dictionary(Of Integer, Vertex), ByRef tempFaces As Dictionary(Of Integer, Face))
        ' Update face vertex ids to the new vertex ids
        For Each pair As KeyValuePair(Of Integer, Face) In tempFaces
            Dim face As Face = pair.Value
            For Each faceVertex As Integer() In face.FaceVertices
                For i As Integer = 0 To faceVertex.Length - 1
                    Dim vertex As Vertex = tempVertices(faceVertex(i))
                    faceVertex(i) = vertex.NewId
                Next i
            Next
        Next
    End Sub

    Sub shiftVIds(ByVal gap As Boolean, ByVal startId As Integer, ByVal diff As Integer, ByRef tempVertices As Dictionary(Of Integer, Vertex), ByRef tempFaces As Dictionary(Of Integer, Face))
        For Each pair As KeyValuePair(Of Integer, Vertex) In tempVertices
            Dim vId As Integer = pair.Key
            Dim vertex As Vertex = pair.Value
            If vId > startId Or gap = False Then
                vertex.NewId = vertex.Id - diff + 1
            End If
        Next
        updateFaceVIds(tempVertices, tempFaces)
        Dim copyTempVertices As New Dictionary(Of Integer, Vertex)
        For Each pair As KeyValuePair(Of Integer, Vertex) In tempVertices
            Dim vertex As Vertex = pair.Value
            vertex.Id = vertex.NewId
            copyTempVertices.Add(vertex.Id, vertex)
        Next
        tempVertices = copyTempVertices
    End Sub

    Sub shiftFIds(ByVal diff As Integer, ByRef tempFaces As Dictionary(Of Integer, Face), ByRef tempGroups As Dictionary(Of String, Group))
        For Each pair As KeyValuePair(Of Integer, Face) In tempFaces
            Dim fId As Integer = pair.Key
            Dim face As Face = pair.Value
            face.NewId = face.Id - diff + 1
        Next

        ' Update group faces ids to the new faces ids
        For Each pair As KeyValuePair(Of String, Group) In tempGroups
            Dim group As Group = pair.Value
            Dim tmpFacesList As New List(Of Integer)
            For Each fId As Integer In group.Faces
                Dim face As Face = tempFaces(fId)
                tmpFacesList.Add(face.NewId)
            Next
            group.Faces = tmpFacesList
        Next

        Dim copyTempFaces As New Dictionary(Of Integer, Face)
        For Each pair As KeyValuePair(Of Integer, Face) In tempFaces
            Dim face As Face = pair.Value
            face.Id = face.NewId
            copyTempFaces.Add(face.Id, face)
        Next
        tempFaces = copyTempFaces
    End Sub

    Private Sub CheckedListBoxModels_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckedListBoxModels.SelectedIndexChanged
        currentModel = CheckedListBoxModels.SelectedItem.ToString()
        selectModelGroups(currentModel)
    End Sub

    Sub selectModelGroups(ByVal modelName As String)
        CheckedListBoxGroups.Items.Clear()
        Dim m As Model = models(modelName)
        For Each pair As KeyValuePair(Of String, Group) In m.Groups
            Dim groupName As String = pair.Value.Name
            CheckedListBoxGroups.Items.Add(groupName)
            If m.SelectedGroups.Contains(groupName) = True Then
                Dim idx As Integer = CheckedListBoxGroups.Items.IndexOf(groupName)
                CheckedListBoxGroups.SetItemChecked(idx, True)
            End If
        Next
    End Sub

    Private Sub CheckedListBoxGroups_ItemCheck(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ItemCheckEventArgs) Handles CheckedListBoxGroups.ItemCheck
        Dim m As Model = models(currentModel)
        Dim group As String = CheckedListBoxGroups.Items(e.Index).ToString()
        If e.NewValue = 1 Then
            If m.SelectedGroups.Contains(group) = False Then 'against duplicate adding
                m.SelectedGroups.Add(group)
            End If
        ElseIf e.NewValue = 0 Then
            If m.SelectedGroups.Contains(group) = True Then 'against duplicate adding
                m.SelectedGroups.Remove(group)
            End If
        End If
    End Sub

    Private Sub selectAllGroupsBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles selectAllGroupsBtn.Click
        If togglesSelectAll(currentModel) = False Then
            togglesSelectAll(currentModel) = True
            For i As Integer = 0 To CheckedListBoxGroups.Items.Count - 1
                CheckedListBoxGroups.SetItemChecked(i, True)
            Next i
        Else
            togglesSelectAll(currentModel) = False
            For i As Integer = 0 To CheckedListBoxGroups.Items.Count - 1
                CheckedListBoxGroups.SetItemChecked(i, False)
            Next i
        End If
    End Sub

    Function writeLoc3d(ByVal l As loc3d) As String
        Return " " & l.x & " " & l.y & " " & l.z
    End Function

    Function writeLoc2d(ByVal l As loc2d) As String
        Return " " & l.x & " " & l.y
    End Function

    Sub saveModel(ByVal pathed As String, ByVal model As Model)
        Dim text As String = vbNewLine
        ToolStripProgressBar1.Value = 0
        ToolStripProgressBar1.Maximum = (model.Vertices.Count * 3) + (model.Faces.Count * 3) + model.Groups.Count

        For Each pair As KeyValuePair(Of Integer, Vertex) In model.Vertices
            Dim vertex As Vertex = pair.Value
            text &= "v" & writeLoc3d(vertex.Pos) & vbNewLine
            ToolStripProgressBar1.Value += 1
        Next
        text &= vbNewLine
        For Each pair As KeyValuePair(Of Integer, Vertex) In model.Vertices
            Dim vertex As Vertex = pair.Value
            text &= "vt" & writeLoc2d(vertex.Tex) & vbNewLine
            ToolStripProgressBar1.Value += 1
        Next
        text &= vbNewLine
        For Each pair As KeyValuePair(Of Integer, Vertex) In model.Vertices
            Dim vertex As Vertex = pair.Value
            text &= "vn" & writeLoc3d(vertex.Norm) & vbNewLine
            ToolStripProgressBar1.Value += 1
        Next
        text &= vbNewLine

        For Each pair As KeyValuePair(Of String, Group) In model.Groups
            Dim group As Group = pair.Value
            text &= "g " & group.Name & vbNewLine
            ToolStripProgressBar1.Value += 1
            For Each faceId As Integer In group.Faces
                Dim face As Face = model.Faces(faceId)
                text &= "f "
                For i As Integer = 0 To face.FaceVertices.Count - 1
                    Dim faceVertex As Integer() = face.FaceVertices(i)
                    For j As Integer = 0 To faceVertex.Length - 1
                        Dim vertex As Integer = faceVertex(j)
                        text &= vertex
                        If j < faceVertex.Length - 1 Then
                            text &= "/"
                        End If
                    Next j
                    If i < face.FaceVertices.Count - 1 Then
                        text &= " "
                    End If
                    ToolStripProgressBar1.Value += 1
                Next i
                text &= vbNewLine
            Next
        Next

        My.Computer.FileSystem.WriteAllText(pathed, text, True)
    End Sub

    Private Sub LoadModelBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LoadModelBtn.Click
        Dim canceled As DialogResult

        canceled = OpenFileDialog1.ShowDialog()
        If canceled = Windows.Forms.DialogResult.Cancel Then
            Exit Sub
        End If
    End Sub

    Private Sub OpenFileDialog1_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk
        If OpenFileDialog1.CheckFileExists() Then
            loadModel(OpenFileDialog1.FileName)
        End If
    End Sub

    Private Sub combineBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles combineBtn.Click
        Dim selectedModelItems As CheckedListBox.CheckedItemCollection = CheckedListBoxModels.CheckedItems
        Dim selectedModels As New List(Of String)
        For Each item As Object In selectedModelItems
            selectedModels.Add(item.ToString())
        Next
        If selectedModels.Count >= 1 Then
            outputModel = combineModels(selectedModels)
        Else
            MessageBox.Show("Error! Need at least one model to combine", "Error", 0, MessageBoxIcon.Error)
        End If

        Dim canceled As DialogResult
        canceled = SaveFileDialog1.ShowDialog()
        If canceled = Windows.Forms.DialogResult.Cancel Then
            Exit Sub
        End If
    End Sub

    Private Sub SaveFileDialog1_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles SaveFileDialog1.FileOk
        statusText.Text = "Saving..."
        Application.DoEvents()
        Dim pathed As String = SaveFileDialog1.FileName

        If pathed.Substring(pathed.Length - 4, 1) = "." Then
            If pathed.Substring(pathed.Length - 3) <> "obj" Then
                pathed = pathed.Remove(pathed.Length - 3)
                pathed &= "obj"
            End If
        ElseIf pathed.Substring(pathed.Length - 5, 1) = "." Then
            pathed = pathed.Remove(pathed.Length - 4)
            pathed &= "obj"
        End If

        If System.IO.File.Exists(pathed) = False Then
            saveModel(pathed, outputModel)
            statusText.Text = "Done"
            Application.DoEvents()
            MessageBox.Show("Successful!", "Save", 0, MessageBoxIcon.Information)
            ToolStripProgressBar1.Value = 0
        Else
            MessageBox.Show("File already exists!", "Error", 0, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        statusText.Text = "Ready"
    End Sub

End Class
