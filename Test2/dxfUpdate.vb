Imports Inventor

Module dxfUpdate

    Sub runDxfUpdate()

        Dim invApp As Inventor.Application = Nothing
        Dim oPartDoc As PartDocument = g_inventorApplication.ActiveDocument
        Dim oDef As PartComponentDefinition = oPartDoc.ComponentDefinition
        Dim oFactory As iPartFactory = oDef.iPartFactory
        Dim oRow As iPartTableRow

        'Check if part is a factory
        If oDef.IsiPartFactory = True Then

            'Go through all rows
            For Each oRow In oFactory.TableRows
                oFactory.DefaultRow = oRow



                'Get part number
                Dim PN As String = oPartDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value.ToString
                'Set data write output format
                Dim oFormat As String = "FLAT PATTERN DXF?AcadVersion=2018&OuterProfileLayer=IV_INTERIOR_PROFILES"
                'Set output path
                Dim oFileName As String = "X:\" & PN & ".dxf"

                'Make a flat pattern if one doesn't exist and refold
                If oDef.HasFlatPattern = False Then
                    Try

                        oDef.Unfold
                        oDef.Flatpattern.ExitEdit

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
            Dim oFormat As String = "FLAT PATTERN DXF?AcadVersion=2000&OuterProfileLayer=IV_INTERIOR_PROFILES"
            Dim oFileName As String = "X:\" & PN & ".dxf"

            If oDef.HasFlatPattern = False Then
                oDef.Unfold
                oDef.flatPattern.ExitEdit
            End If

            If g_inventorApplication.ErrorManager.HasErrors = True Then
                MsgBox("Broke")
                Return
            End If

            oDef.DataIO.WriteDataToFile(oFormat, oFileName)

        End If

    End Sub

End Module
