Imports Inventor

Module dxfUpdate

    Sub RunDxfUpdate(ThisApplication As Inventor.Application)
        Dim oPartDoc As PartDocument = ThisApplication.ActiveDocument
        Dim oDef As SheetMetalComponentDefinition = oPartDoc.ComponentDefinition
        Dim oFactory As iPartFactory = oDef.iPartFactory
        Dim oRow As iPartTableRow

        'Check if part is a factory
        If oDef.IsiPartFactory = True Then

            'Go through all rows
            For Each oRow In oFactory.TableRows
                oFactory.DefaultRow = oRow

                Dim partMaterial As String = oPartDoc.PropertySets.Item("Design Tracking Properties").Item("Material").Value.ToString
                MsgBox(partMaterial)

                Dim partAppearance As String = oPartDoc.ActiveAppearance.Name.ToString
                MsgBox(partAppearance)

                If oDef.FlatPattern.FlatBendResults.Count = 0 And (oDef.FlatPattern.RangeBox.MaxPoint.Z - oDef.FlatPattern.RangeBox.MinPoint.Z) > (oDef.Thickness.Value + 0.003) Then
                    MsgBox("Invalid")
                Else
                    MsgBox("Valid")

                End If

                'Get part number
                Dim PN As String = oPartDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value.ToString
                'Set data write output format
                Dim oFormat As String = "FLAT PATTERN DXF?AcadVersion=2018" _
                    + "&BendDownLayer=DOWN&BendDownLayerColor=255;0;0" _
                    + "&BendUpLayer=UP&BendUpLayerColor=255;0;0" _
                    + "&OuterProfileLayer=OUTER&OuterProfileLayerColor=0;0;255" _
                    + "&InteriorProfilesLayer=INNER&InteriorProfilesLayerColor=0;0;0" _
                    + "&ArcCentersLayer=POINT&ArcCentersLayerColor=255;0;255" _
                    + "&TangentLayer=RADIUS&TangentLayerColor=255;255;0"
                'Set output path
                Dim oFileName As String = "X:\" & PN & ".dxf"

                'Make a flat pattern if one doesn't exist and refold

                If oDef.HasFlatPattern = False Then
                    Try

                        oDef.Unfold()

                        oDef.FlatPattern.ExitEdit()

                    Catch ex As Exception

                        MsgBox("Failed to create flat pattern")
                        Return

                    End Try

                End If

                'Create dxf
                Try

                    oDef.DataIO.WriteDataToFile(oFormat, oFileName)

                Catch ex As Exception

                    MsgBox("DXF failed to generate")
                    Return

                End Try


            Next

            'Reset to first member
            oRow = oFactory.TableRows.Item(1)
            oFactory.DefaultRow = oRow
            Dim Total As Integer = oFactory.TableRows.Count
            MsgBox("Created " & Total)
        Else

            'All same as above, for non-iparts

            Dim PN As String = oPartDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value.ToString
            Dim oFormat As String = "FLAT PATTERN DXF?AcadVersion=2018" _
                    + "&BendDownLayer=DOWN&BendDownLayerColor=255;0;0" _
                    + "&BendUpLayer=UP&BendUpLayerColor=255;0;0" _
                    + "&OuterProfileLayer=OUTER&OuterProfileLayerColor=0;0;255" _
                    + "&InteriorProfilesLayer=INNER&InteriorProfilesLayerColor=0;0;0" _
                    + "&ArcCentersLayer=POINT&ArcCentersLayerColor=255;0;255" _
                    + "&TangentLayer=RADIUS&TangentLayerColor=255;255;0"

            Dim oFileName As String = "X:\" & PN & ".dxf"

            If oDef.HasFlatPattern = False Then
                oDef.Unfold()
                oDef.FlatPattern.ExitEdit()
            End If

            Try
                oDef.DataIO.WriteDataToFile(oFormat, oFileName)
            Catch ex As Exception
                MsgBox("Check to see if you are connected to the X Drive and try again")
                Return
            End Try
        End If

    End Sub

End Module