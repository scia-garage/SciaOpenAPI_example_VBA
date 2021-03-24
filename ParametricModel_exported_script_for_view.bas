Attribute VB_Name = "Module1"
Option Explicit
Private Sub DeleteTemp(TempFolder As String)

    Dim FSO As Scripting.FileSystemObject
    Set FSO = New Scripting.FileSystemObject

    If FSO.FolderExists(TempFolder) Then
        Call FSO.DeleteFolder(TempFolder, True)
    End If


End Sub

Public Function OpenApiExample() As Integer
 Dim TempFolder As String
 TempFolder = Sheets("Input").Cells(12, 2).Value
 DeleteTemp (TempFolder)
 
  Dim SENpath As String
  SENpath = Sheets("Input").Cells(10, 2).Value
  
  Dim templ As String
  templ = Sheets("Input").Cells(11, 2).Value

  
'  Debug.Print SENpath
'  Debug.Print templ
  
  
  Dim env As New SCIA_OpenAPI.Environment
 
  'Setting of environment
  Call env.Init(SENpath, ".\tmp", "1.0.0.0")
  
  'Start of SCIA Engineer
  Dim openedSE As Boolean
  openedSE = env.RunSCIAEngineer(GuiMode_ShowWindowShow)
  If Not openedSE Then
  Exit Function
  End If
 ' Debug.Print "SEn started"
  'Open empty project
  Dim proj As EsaProject
  Set proj = env.OpenProject(templ)
  If proj = Null Then
  Exit Function
  End If
  
  'Debug.Print "Template opened"
  
  'Create materials in local ADM
  Dim comatid As New ApiGuid
  Call comatid.SetFromString(Get_NewGUID())
  Dim conmat As SCIA_OpenAPI.Material
  Set conmat = New SCIA_OpenAPI.Material
  Set conmat.ID = comatid
  conmat.Name = Sheets("Input").Cells(4, 2).Value
  conmat.Type = 0
  conmat.Quality = Sheets("Input").Cells(4, 2).Value
  
  Call proj.Model.CreateMaterial(conmat)
  
  Dim stmatid As New ApiGuid
  Call stmatid.SetFromString(Get_NewGUID())
  Dim stmat As SCIA_OpenAPI.Material
  Set stmat = New SCIA_OpenAPI.Material
  Set stmat.ID = stmatid
  stmat.Name = Sheets("Input").Cells(5, 2)
  stmat.Type = 1
  stmat.Quality = Sheets("Input").Cells(5, 2).Value
  Call proj.Model.CreateMaterial(stmat)
    
  'Cross-sections in local ADM
   Dim cssid As New ApiGuid
  Call cssid.SetFromString(Get_NewGUID())
  Dim css As New SCIA_OpenAPI.CrossSectionManufactured
  Set css.ID = cssid
  css.Name = "css"
  css.Profile = Sheets("Input").Cells(6, 2).Value
  css.FormCode = 1
  css.DescriptionId = 0
  Set css.Material = stmat.ID
  Call proj.Model.CreateCrossSection(css)
  
  Dim n1id As New ApiGuid
  Call n1id.SetFromString(Get_NewGUID())
  Dim n2id As New ApiGuid
  Call n2id.SetFromString(Get_NewGUID())
  Dim n3id As New ApiGuid
  Call n3id.SetFromString(Get_NewGUID())
  Dim n4id As New ApiGuid
  Call n4id.SetFromString(Get_NewGUID())
  Dim n5id As New ApiGuid
  Call n5id.SetFromString(Get_NewGUID())
  Dim n6id As New ApiGuid
  Call n6id.SetFromString(Get_NewGUID())
  Dim n7id As New ApiGuid
  Call n7id.SetFromString(Get_NewGUID())
  Dim n8id As New ApiGuid
  Call n8id.SetFromString(Get_NewGUID())
  
  
  Dim n1 As SCIA_OpenAPI.StructNode
  Set n1 = New SCIA_OpenAPI.StructNode
  Set n1.ID = n1id
  n1.Name = "N1"
  n1.X = 0#
  n1.Y = 0#
  n1.Z = 0#
  Call proj.Model.CreateNode(n1)
  
  Dim n2 As New StructNode
  Set n2 = New SCIA_OpenAPI.StructNode
  Set n2.ID = n2id
  n2.Name = "N2"
  n2.X = Sheets("Input").Cells(1, 2) 'a
  n2.Y = 0
  n2.Z = 0
  Call proj.Model.CreateNode(n2)
  
  Dim n3 As SCIA_OpenAPI.StructNode
  Set n3 = New SCIA_OpenAPI.StructNode
  Set n3.ID = n3id
  n3.Name = "N3"
  n3.X = Sheets("Input").Cells(1, 2) 'a
  n3.Y = Sheets("Input").Cells(2, 2) 'b
  n3.Z = 0#
  Call proj.Model.CreateNode(n3)
  
  Dim n4 As New StructNode
  Set n4 = New SCIA_OpenAPI.StructNode
  Set n4.ID = n4id
  n4.Name = "N4"
  n4.X = 0
  n4.Y = Sheets("Input").Cells(2, 2) 'b
  n4.Z = 0
  Call proj.Model.CreateNode(n4)
  
  Dim n5 As SCIA_OpenAPI.StructNode
  Set n5 = New SCIA_OpenAPI.StructNode
  Set n5.ID = n5id
  n5.Name = "N5"
  n5.X = 0#
  n5.Y = 0#
  n5.Z = Sheets("Input").Cells(3, 2) 'c
  Call proj.Model.CreateNode(n5)
  
  Dim n6 As New StructNode
  Set n6 = New SCIA_OpenAPI.StructNode
  Set n6.ID = n6id
  n6.Name = "N6"
  n6.X = Sheets("Input").Cells(1, 2) 'a
  n6.Y = 0 'b
  n6.Z = Sheets("Input").Cells(3, 2) 'c
  Call proj.Model.CreateNode(n6)
  
  Dim n7 As SCIA_OpenAPI.StructNode
  Set n7 = New SCIA_OpenAPI.StructNode
  Set n7.ID = n7id
  n7.Name = "N7"
  n7.X = Sheets("Input").Cells(1, 2) 'a
  n7.Y = Sheets("Input").Cells(2, 2) 'b
  n7.Z = Sheets("Input").Cells(3, 2) 'c
  Call proj.Model.CreateNode(n7)
  
  Dim n8 As New StructNode
  Set n8 = New SCIA_OpenAPI.StructNode
  Set n8.ID = n8id
  n8.Name = "N8"
  n8.X = 0
  n8.Y = Sheets("Input").Cells(2, 2) 'b
  n8.Z = Sheets("Input").Cells(3, 2) 'c
  Call proj.Model.CreateNode(n8)
  
  
  Dim nodesb1 As SCIA_OpenAPI.ApiGuidArr
  Set nodesb1 = New SCIA_OpenAPI.ApiGuidArr
  Call nodesb1.Init(2)
  Call nodesb1.SetElementFromGuidString(0, n1id.ReturnString)
  Call nodesb1.SetElementFromGuidString(1, n5id.ReturnString)
  Dim b1id As New ApiGuid
  Call b1id.SetFromString(Get_NewGUID())
  Dim b1 As New SCIA_OpenAPI.Beam
  Set b1.ID = b1id
  b1.Name = "beam1"
  Set b1.css = css.ID
  Set b1.nodes = nodesb1
  Call proj.Model.CreateBeam(b1)
  
  Dim nodesb2 As SCIA_OpenAPI.ApiGuidArr
  Set nodesb2 = New SCIA_OpenAPI.ApiGuidArr
  Call nodesb2.Init(2)
  Call nodesb2.SetElementFromGuidString(0, n2id.ReturnString)
  Call nodesb2.SetElementFromGuidString(1, n6id.ReturnString)
  Dim b2id As New ApiGuid
  Call b2id.SetFromString(Get_NewGUID())
  Dim b2 As New SCIA_OpenAPI.Beam
  Set b2.ID = b2id
  b2.Name = "beam2"
  Set b2.css = css.ID
  Set b2.nodes = nodesb2
  Call proj.Model.CreateBeam(b2)
    
  
  
  Dim nodesb3 As SCIA_OpenAPI.ApiGuidArr
  Set nodesb3 = New SCIA_OpenAPI.ApiGuidArr
  Call nodesb3.Init(2)
  Call nodesb3.SetElementFromGuidString(0, n3id.ReturnString)
  Call nodesb3.SetElementFromGuidString(1, n7id.ReturnString)
  Dim b3id As New ApiGuid
  Call b3id.SetFromString(Get_NewGUID())
  Dim b3 As New SCIA_OpenAPI.Beam
  Set b3.ID = b3id
  b3.Name = "beam3"
  Set b3.css = css.ID
  Set b3.nodes = nodesb3
  Call proj.Model.CreateBeam(b3)
  
  Dim nodesb4 As SCIA_OpenAPI.ApiGuidArr
  Set nodesb4 = New SCIA_OpenAPI.ApiGuidArr
  Call nodesb4.Init(2)
  Call nodesb4.SetElementFromGuidString(0, n4id.ReturnString)
  Call nodesb4.SetElementFromGuidString(1, n8id.ReturnString)
  Dim b4id As New ApiGuid
  Call b4id.SetFromString(Get_NewGUID())
  Dim b4 As New SCIA_OpenAPI.Beam
  Set b4.ID = b4id
  b4.Name = "beam4"
  Set b4.css = css.ID
  Set b4.nodes = nodesb4
  Call proj.Model.CreateBeam(b4)
   
  
  Dim sup1id As New ApiGuid
  Call sup1id.SetFromString(Get_NewGUID())
  Dim sup1 As New SCIA_OpenAPI.PointSupport
  Set sup1.ID = sup1id
  sup1.Name = "PS1"
  Set sup1.Member = n1id
  sup1.ConstraintRx = eConstraintType_Free
  sup1.ConstraintRy = eConstraintType_Free
  sup1.ConstraintRz = eConstraintType_Free
  
  Call proj.Model.CreatePointSupport(sup1)
  
  Dim sup2id As New ApiGuid
  Call sup2id.SetFromString(Get_NewGUID())
  Dim sup2 As New SCIA_OpenAPI.PointSupport
  Set sup2.ID = sup2id
  sup2.Name = "PS2"
  Set sup2.Member = n2id
  Call proj.Model.CreatePointSupport(sup2)
  
  Dim sup3id As New ApiGuid
  Call sup3id.SetFromString(Get_NewGUID())
  Dim sup3 As New SCIA_OpenAPI.PointSupport
  Set sup3.ID = sup3id
  sup3.Name = "PS3"
  Set sup3.Member = n3id
  Call proj.Model.CreatePointSupport(sup3)
  
  Dim sup4 As New SCIA_OpenAPI.PointSupport
  Dim sup4id As New ApiGuid
  Call sup4id.SetFromString(Get_NewGUID())
  Set sup4.ID = sup4id
  sup4.Name = "PS4"
  Set sup4.Member = n4id
  Call proj.Model.CreatePointSupport(sup4)
  
  
  
  Dim nodess1 As SCIA_OpenAPI.ApiGuidArr
  Set nodess1 = New SCIA_OpenAPI.ApiGuidArr
  Call nodess1.Init(4)
  Call nodess1.SetElementFromGuidString(0, n5id.ReturnString)
  Call nodess1.SetElementFromGuidString(1, n6id.ReturnString)
  Call nodess1.SetElementFromGuidString(2, n7id.ReturnString)
  Call nodess1.SetElementFromGuidString(3, n8id.ReturnString)
  Dim s1 As New SCIA_OpenAPI.Slab
  Dim s1id As New ApiGuid
  Call s1id.SetFromString(Get_NewGUID())
  Set s1.ID = s1id
  s1.Name = "S1"
  Set s1.nodes = nodess1
  Set s1.MaterialId = conmat.ID
  s1.Thickness = Sheets("Input").Cells(7, 2).Value
  s1.Type = 0
  Call proj.Model.CreateSlab(s1)
  
  Dim lSupport As New SCIA_OpenAPI.LineSupport
  Dim lSupportid As New ApiGuid
  Call lSupportid.SetFromString(Get_NewGUID())
  Set lSupport.ID = lSupportid
  lSupport.Name = "LineSupport"
  Set lSupport.Member = b1.ID
  lSupport.ConstraintRx = SCIA_OpenAPI.eConstraintType.eConstraintType_Free
  lSupport.ConstraintRy = SCIA_OpenAPI.eConstraintType.eConstraintType_Free
  lSupport.ConstraintRz = SCIA_OpenAPI.eConstraintType.eConstraintType_Free
  Call proj.Model.CreateLineSupport(lSupport)
 
  
  Dim lg1id As New ApiGuid
  Call lg1id.SetFromString(Get_NewGUID())
  Dim lg1 As New SCIA_OpenAPI.LoadGroup
  Set lg1.ID = lg1id
  lg1.Name = "LoadGroup1"
  lg1.Type = 0
   Call proj.Model.CreateLoadGroup(lg1)
  
  Dim lc1id As New ApiGuid
  Call lc1id.SetFromString(Get_NewGUID())
  Dim lc1 As New SCIA_OpenAPI.LoadCase
  Set lc1.ID = lc1id
  lc1.Name = "LoadCase1"
  Set lc1.LoadGroupId = lg1id
  lc1.ActionType = 0
  lc1.LoadCaseType = 1
  Call proj.Model.CreateLoadCase(lc1)
  
  ' Combination
'  Dim combinationItems As New Collection
'  Dim CI1 As New CombinationItem
'  CI1.Coefficient = 1.5
'  Set CI1.LoadCase = lc1id
'  combinationItems.Add (CI1)
'  Dim C1 As New SCIA_OpenAPI.Combination
'  Dim C1id As New ApiGuid
'  Call C1id.SetFromString(Get_NewGUID())
'  Set C1.ID = C1id
'  C1.Name = "C1"
'  Call C1.SetCombinationContentVBA(combinationItems)
'  C1.NationalStandard = eLoadCaseCombinationStandard.eLoadCaseCombinationStandard_EnUlsSetB
'  Call proj.Model.CreateCombination(C1)
  
  Dim sl1id As New ApiGuid
  Call sl1id.SetFromString(Get_NewGUID())
  Dim sl1 As New SCIA_OpenAPI.SurfaceLoad
  Set sl1.ID = sl1id
  sl1.Name = "SL1"
  sl1.Direction = 2
  Set sl1.LoadCaseId = lc1id
  Set sl1.Member2DId = s1.ID
  sl1.Value = Sheets("Input").Cells(8, 2).Value
  Call proj.Model.CreateSurfaceLoad(sl1)
  
  Dim lLoad As New SCIA_OpenAPI.LineLoadOnBeam
  Dim lLoadtid As New ApiGuid
  Call lLoadtid.SetFromString(Get_NewGUID())
  Set lLoad.ID = lLoadtid
  lLoad.Name = "LineLoad"
  Set lLoad.Member = b1.ID
  Set lLoad.LoadCase = lc1.ID
  lLoad.Value1 = -12500
  lLoad.Value2 = -12500
  lLoad.Direction = SCIA_OpenAPI.eDirection.eDirection_X
  Call proj.Model.CreateLineLoad(lLoad)

  
  Dim lLoadEdge As New SCIA_OpenAPI.LineLoadOnSlabEdge
  Dim lLoadEdgeid As New ApiGuid
  Call lLoadEdgeid.SetFromString(Get_NewGUID())
  Set lLoadEdge.ID = lLoadEdgeid
  lLoadEdge.Name = "LineLoadEdge"
  Set lLoadEdge.Member = s1.ID
  lLoadEdge.EdgeIndex = 0
  Set lLoadEdge.LoadCase = lc1.ID
  lLoadEdge.Value1 = -12500
  lLoadEdge.Value2 = -12500
  lLoadEdge.Direction = SCIA_OpenAPI.eDirection.eDirection_X
  Call proj.Model.CreateLineLoad_2(lLoadEdge)
  
  Call proj.Model.RefreshModel_ToSCIAEngineer
  
  
  Call proj.RunCalculation

  Dim rapi As ResultsAPI
  Set rapi = proj.Model.InitializeResultsAPI()
 

  Dim keyIntForcesB1 As New ResultKey
  keyIntForcesB1.EntityType = eDsElementType_eDsElementType_Beam
  keyIntForcesB1.EntityName = "beam1"
  keyIntForcesB1.CaseType = eDsElementType_eDsElementType_LoadCase
  Set keyIntForcesB1.CaseId = lc1id
  keyIntForcesB1.Dimension = eDimension_eDim_1D
  keyIntForcesB1.ResultType = eResultType_eFemBeamInnerForces
  keyIntForcesB1.CoordSystem = eCoordSystem_eCoordSys_Local

  Dim rintf As Result
'  Set rintf = rapi.LoadResult(keyIntForcesB1)
  
  'Dim keyIntForcesB1Combi As New ResultKey
  'keyIntForcesB1Combi.EntityType = eDsElementType_eDsElementType_Beam
  'keyIntForcesB1Combi.EntityName = "beam1"
  'keyIntForcesB1Combi.CaseType = eDsElementType_eDsElementType_Combination
  'Set keyIntForcesB1Combi.CaseId = lc1id
  'keyIntForcesB1Combi.Dimension = eDimension_eDim_1D
  'keyIntForcesB1Combi.ResultType = eResultType_eFemBeamInnerForces
  'keyIntForcesB1Combi.CoordSystem = eCoordSystem_eCoordSys_Local

'  Dim rintfCombi As Result
'  Set rintfCombi = rapi.LoadResult(keyIntForcesB1Combi)
  
'Dim resulttable As String
'  resulttable = rintf.GetTextOutput()
'  Debug.Print (resulttable)
'
  
  Dim i, j As Integer
  i = 0
  j = 0
  
' Do While i < rintf.GetMeshElementCount()
'  Sheets("Result").Cells(i + 3, j + 1) = i + 1
'    Do While j < rintf.GetMagnitudesCount()
'    Sheets("Result").Cells(i + 3, j + 2) = rintf.GetValue(j, i)
'    j = j + 1
'  Loop
'  i = i + 1
'  j = 0
'  Loop


  Dim keySlabDef As New ResultKey
  keySlabDef.EntityType = eDsElementType_eDsElementType_Slab
  keySlabDef.EntityName = "S1"
  keySlabDef.CaseType = eDsElementType_eDsElementType_LoadCase
  Set keySlabDef.CaseId = lc1id
  keySlabDef.Dimension = eDimension_eDim_2D
  keySlabDef.ResultType = eResultType_eFemDeformations
  keySlabDef.CoordSystem = eCoordSystem_eCoordSys_Local

  Dim defSlabRes As Result
'  Set defSlabRes = rapi.LoadResult(keySlabDef)

i = 0
j = 0

'Do While i < defSlabRes.GetMeshElementCount()
'    Sheets("Result").Cells(i + 18, j + 1) = i + 1
'    Do While j < defSlabRes.GetMagnitudesCount()
'      Sheets("Result").Cells(i + 18, j + 2) = defSlabRes.GetValue(j, i)
'    j = j + 1
'  Loop
'  i = i + 1
'  j = 0
'  Loop

    'Dim resulttable As String
    'resulttable = defSlabRes.GetTextOutput()
    'Debug.Print (resulttable)


   Call proj.CloseProject(SaveMode_SaveChangesNo)
   Call env.Dispose
   MsgBox ("Scia Engineer run finished. See result tab.")

'  OpenApiExample = 0
  
Error:
  Debug.Print "Chyba"
  Call env.Dispose
  End Function
