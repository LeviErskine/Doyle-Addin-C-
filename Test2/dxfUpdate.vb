Imports Inventor

Module dxfUpdate

    Sub RunDxfUpdate(ThisApplication As Inventor.Application)
        Dim oPartDoc As PartDocument = ThisApplication.ActiveDocument
        Dim oDef As SheetMetalComponentDefinition = oPartDoc.ComponentDefinition
        Dim oFactory As iPartFactory = oDef.iPartFactory
        Dim oRow As iPartTableRow
        Dim failedExports As New List(Of String)

        'Check if part is a factory
        If oDef.IsiPartFactory = True Then

            'Go through all rows
            For Each oRow In oFactory.TableRows
                oFactory.DefaultRow = oRow
                'Debug
                ' MsgBox(ThisApplication.ActiveDocument._SickNodesCount.ToString, MsgBoxStyle.OkOnly, "Sick Nodes")
                ' MsgBox(ThisApplication.ActiveDocument._ComatoseNodesCount.ToString, MsgBoxStyle.OkOnly, "Comatose Nodes")

                Dim memberDef As SheetMetalComponentDefinition = TryCast(oPartDoc.ComponentDefinition, SheetMetalComponentDefinition)
                Dim PN As String = oPartDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value.ToString
                Dim oFileName As String = "X:\" & PN & ".dxf"

                ' Check for comatose nodes
                If oPartDoc._ComatoseNodesCount > 0 Then
                    failedExports.Add("Broken feature detected for: " & PN)
                    Continue For
                End If

                If oPartDoc._SickNodesCount > 0 Then
                    failedExports.Add("Potentally troublesome feature detected for: " & PN)
                    Continue For
                End If

                ' Create flat pattern if missing
                If Not memberDef.HasFlatPattern Then
                    Try
                        memberDef.Unfold()
                        memberDef.FlatPattern.ExitEdit()
                    Catch ex As Exception
                        failedExports.Add("Failed to create flat pattern for: " & PN)
                        Continue For
                    End Try
                End If

                ' Validate flat pattern
                If memberDef.FlatPattern.FlatBendResults.Count = 0 And
                    (memberDef.FlatPattern.RangeBox.MaxPoint.Z - memberDef.FlatPattern.RangeBox.MinPoint.Z) > (memberDef.Thickness.Value + 0.003) Then
                    failedExports.Add("Invalid flat pattern for: " & PN)
                    Continue For
                End If

                ' Prepare output path and format
                Dim oFormat As String = "FLAT PATTERN DXF?AcadVersion=2018" _
                    & "&BendDownLayer=DOWN&BendDownLayerColor=255;0;0" _
                    & "&BendUpLayer=UP&BendUpLayerColor=255;0;0" _
                    & "&OuterProfileLayer=OUTER&OuterProfileLayerColor=0;0;255" _
                    & "&InteriorProfilesLayer=INNER&InteriorProfilesLayerColor=0;0;0" _
                    & "&ArcCentersLayer=POINT&ArcCentersLayerColor=255;0;255" _
                    & "&TangentLayer=RADIUS&TangentLayerColor=255;255;0"

                ' Export DXF
                Try
                    memberDef.DataIO.WriteDataToFile(oFormat, oFileName)
                Catch ex As Exception
                    failedExports.Add("DXF failed to generate for: " & PN)
                    Continue For
                End Try
            Next

            'Reset to first member
            oRow = oFactory.TableRows.Item(1)
            oFactory.DefaultRow = oRow
            Dim Total As Integer = oFactory.TableRows.Count

            If failedExports.Count > 0 Then
                MsgBox(failedExports.Count & " Members have errors and were skipped." & vbCrLf & String.Join(vbCrLf, failedExports))
            Else
                MsgBox("Created " & Total & " DXFs. All exports succeeded.")
            End If
        Else
            'Debug
            ' MsgBox(ThisApplication.ActiveDocument._SickNodesCount.ToString, MsgBoxStyle.OkOnly, "Sick Nodes")
            ' MsgBox(ThisApplication.ActiveDocument._ComatoseNodesCount.ToString, MsgBoxStyle.OkOnly, "Comatose Nodes")
            'All same as above, for non-iparts
            Dim PN As String = oPartDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value.ToString
            Dim oFileName As String = "X:\" & PN & ".dxf"
            Dim oFormat As String = "FLAT PATTERN DXF?AcadVersion=2018" _
                        + "&BendDownLayer=DOWN&BendDownLayerColor=255;0;0" _
                        + "&BendUpLayer=UP&BendUpLayerColor=255;0;0" _
                        + "&OuterProfileLayer=OUTER&OuterProfileLayerColor=0;0;255" _
                        + "&InteriorProfilesLayer=INNER&InteriorProfilesLayerColor=0;0;0" _
                        + "&ArcCentersLayer=POINT&ArcCentersLayerColor=255;0;255" _
                        + "&TangentLayer=RADIUS&TangentLayerColor=255;255;0"

            ' Skip export if comatose nodes exist
            If oPartDoc._ComatoseNodesCount > 0 Or oPartDoc._SickNodesCount > 0 Then
                MsgBox(oPartDoc.DisplayName & " has errors, fix before export." & vbCrLf & String.Join(vbCrLf, failedExports))
                Return
            End If

            If oDef.HasFlatPattern = False Then
                Try
                    oDef.Unfold()
                    oDef.FlatPattern.ExitEdit()
                Catch ex As Exception
                    MsgBox("Failed to create flat pattern", MsgBoxStyle.OkOnly, "Error")
                    Return
                End Try
            End If

            Try
                oDef.DataIO.WriteDataToFile(oFormat, oFileName)
            Catch ex As Exception
                MsgBox("DXF failed to generate. Check connection to X drive", MsgBoxStyle.OkOnly, "Error")
                Return
            End Try
        End If
    End Sub

End Module