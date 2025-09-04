Imports System.IO
Imports Doyle_Addin.Options
Imports Inventor

Module DxfUpdate
	Sub RunDxfUpdate(thisApplication As Application)
		Dim oPartDoc As PartDocument = thisApplication.ActiveDocument
		Dim oDef As SheetMetalComponentDefinition = oPartDoc.ComponentDefinition
		Dim oFactory As iPartFactory = oDef.iPartFactory
		Dim failedExports As New List(Of String)
		Dim oDoc As Documents = thisApplication.Documents
		Dim pn As String = oPartDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value.ToString
		Dim oFileName As String = UserOptions.Load.DxfExportLocation & pn & ".dxf"


		'Check if a part is a factory
		If oDef.IsiPartFactory = True Then
			Dim total As Integer = oFactory.TableRows.Count
			oDoc.CloseAll(True)
			oPartDoc.ReleaseReference()

			' Create a flat pattern if missing
			If Not oDef.HasFlatPattern Then
				Try
					oDef.Unfold()
					oDef.FlatPattern.ExitEdit()
				Catch ex As Exception
					failedExports.Add("Failed to create flat pattern for: " & pn)
				End Try
			End If
			If Directory.Exists(oFactory.MemberCacheDir) = False Then
				Directory.CreateDirectory(oFactory.MemberCacheDir)
			End If
			Dim partFiles As String() = Directory.GetFiles(oFactory.MemberCacheDir)

			If total > partFiles.Length Then
				Dim result As MsgBoxResult =
					    MsgBox(
						    "Warning: The factory has " & total & " members, but " & partFiles.Length &
						    " files were found in the folder. Generate files?.",
						    MsgBoxStyle.YesNo,
						    "Missing Members")
				If result = MsgBoxResult.Yes Then
					oDoc.CloseAll(True)
					oPartDoc.ReleaseReference()
					For Each iPartTableRow In oFactory.TableRows
						oFactory.CreateMember(iPartTableRow)
					Next
				Else
					Exit Sub
				End If
			End If

			For Each filepath As String In Directory.GetFiles(oFactory.MemberCacheDir)

				Dim openedDoc As PartDocument = oDoc.Open(filepath, True)
				Dim memberDef = TryCast(openedDoc.ComponentDefinition, SheetMetalComponentDefinition)
				Dim partnumber As String =
					    openedDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value.ToString
				Dim oiFileName As String = UserOptions.Load.DxfExportLocation & partnumber & ".dxf"


				' Validate flat pattern
				If _
					memberDef.FlatPattern.FlatBendResults.Count = 0 And
					(memberDef.FlatPattern.RangeBox.MaxPoint.Z - memberDef.FlatPattern.RangeBox.MinPoint.Z) >
					(memberDef.Thickness.Value + 0.003) Then
					failedExports.Add("Invalid flat pattern for: " & partnumber)
					Continue For
				End If

				' Prepare output path and format
				Const oFormat As String = "FLAT PATTERN DXF?AcadVersion=2018" & "&BendDownLayer=DOWN&BendDownLayerColor=255;0;0" &
				                          "&BendUpLayer=UP&BendUpLayerColor=255;0;0" &
				                          "&OuterProfileLayer=OUTER&OuterProfileLayerColor=0;0;255" &
				                          "&InteriorProfilesLayer=INNER&InteriorProfilesLayerColor=0;0;0" &
				                          "&ArcCentersLayer=POINT&ArcCentersLayerColor=255;0;255" &
				                          "&TangentLayer=RADIUS&TangentLayerColor=255;255;0"

				' Export DXF
				Try
					memberDef.DataIO.WriteDataToFile(oFormat, oiFileName)
					If Not openedDoc Is Nothing Then
						openedDoc.Close(True)
					End If
				Catch ex As Exception
					failedExports.Add("DXF failed to generate for: " & partnumber)
					Continue For
				End Try
			Next

			If failedExports.Count > 0 Then
				MsgBox(failedExports.Count & " Members have errors and were skipped." & vbCrLf & String.Join(vbCrLf, failedExports))
			Else
				MsgBox("Created " & total & " DXFs. All exports succeeded.")
			End If

		Else
			'Debug
			' MsgBox(ThisApplication.ActiveDocument._SickNodesCount.ToString, MsgBoxStyle.OkOnly, "Sick Nodes")
			' MsgBox(ThisApplication.ActiveDocument._ComatoseNodesCount.ToString, MsgBoxStyle.OkOnly, "Comatose Nodes")
			'All the same as above, for non-iparts
			Const oFormat As String = "FLAT PATTERN DXF?AcadVersion=2018" + "&BendDownLayer=DOWN&BendDownLayerColor=255;0;0" +
			                          "&BendUpLayer=UP&BendUpLayerColor=255;0;0" +
			                          "&OuterProfileLayer=OUTER&OuterProfileLayerColor=0;0;255" +
			                          "&InteriorProfilesLayer=INNER&InteriorProfilesLayerColor=0;0;0" +
			                          "&ArcCentersLayer=POINT&ArcCentersLayerColor=255;0;255" +
			                          "&TangentLayer=RADIUS&TangentLayerColor=255;255;0"

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
				MsgBox(oPartDoc.DisplayName & " exported successfully.", MsgBoxStyle.Information, "Success")
			Catch ex As Exception
				MsgBox("DXF failed to generate. Check connection to X drive" & vbCrLf & "Error: " & ex.Message,
				       MsgBoxStyle.Critical,
				       "Error")
			End Try
		End If
	End Sub
End Module