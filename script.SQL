' Created: 14/11/2018
' Modified: 14/11/2018
' Model: Microsoft Access 2007-2010
' Database: MS Access 2007-2010

'=======================================================
'=== Microsoft Access 2000/2002/2003 database creation method
'===
'=== 1. Create a new database in the Microsoft Access
'=== 2. Create a new module
'=== 3. Copy the TDM3 output SQL script into the new Microsoft Access module
'=== 4. Select from main menu "Tools" item "References..." and check
'===    the "Microsoft ActiveX Data Objects 2.x Library"
'===    and "Microsoft ADO Ext. 2.x for DDL and Security"
'===    and "Microsoft DAO 3.6 Object Library"
'=== 5. Place your mouse cursor somewhere in the main procedure Main()
'=== 6. Run the module code (Click the "Run Sub/UserForm" button or press F5)
'=======================================================

Public con As New ADODB.Connection
Public cat As New ADOX.Catalog
Public tbl As ADOX.Table
Public idx As ADOX.Index
Public dbs As DAO.Database

Sub Main()

Set con = CurrentProject.Connection
cat.ActiveConnection = con
Set dbs = CurrentDb()

On Error GoTo ErrorHandler
  Call CreateTables
  Call CreatePrimaryKeys
  Call CreateIndexes
  Call CreateAlterKeys
  Call CreateRelations
  Call CreateRelationsDAO
  Call CreateQueries

  MsgBox "Script successfully processed.", vbInformation
  Exit Sub

ErrorHandler:
  Select Case Err.Number
    Case -2147217857 'DAO: 3010
      MsgBox "Table " & tbl.Name & " already exist!", vbInformation
      Err.Clear
    Case -2147217868 'DAO: 3284
      MsgBox "Index " & idx.Name & " for table " & tbl.Name & " already exist!", vbInformation
      Err.Clear
    Case Else
      MsgBox Err.Description, vbCritical
  End Select

End Sub

' Create tables
'===============

Sub CreateTables()

Call CreateTable1 'usuario
Call CreateTable2 'nivel
Call CreateTable3 'color
Call CreateTable4 'nivel_color
Call CreateTable5 'panel_participacion
Call CreateTable6 'estatus

End Sub

'=== Create table usuario ========

Sub CreateTable1()
Set tbl = New ADOX.Table

tbl.Name = "usuario"
cat.Tables.Append tbl

Call AddFieldToTable("usuario", "idUsuario", adInteger, 0, 0, "Autoincrement", "", "", "", "", "", "")
Call AddFieldToTable("usuario", "nombre", adVarWChar, 50, 0, "", "TRUE", "", "", "", "", "")
Call AddFieldToTable("usuario", "aPaterno", adVarWChar, 50, 0, "", "TRUE", "", "", "", "", "")
Call AddFieldToTable("usuario", "aMAterno", adVarWChar, 50, 0, "", "TRUE", "", "", "", "", "")
Call AddFieldToTable("usuario", "telefonoMovil", adVarWChar, 50, 0, "", "TRUE", "", "", "", "", "")
Call AddFieldToTable("usuario", "telefonoFijo", adVarWChar, 50, 0, "", "TRUE", "", "", "", "", "")
Call AddFieldToTable("usuario", "direccion", adVarWChar, 255, 0, "", "TRUE", "", "", "", "", "")
Call AddFieldToTable("usuario", "fechaIngreso", adDate, 0, 0, "", "TRUE", "", "", "", "", "")
Call AddFieldToTable("usuario", "idEstatus", adSmallInt, 0, 0, "", "FALSE", "", "", "", "", "")


dbs.TableDefs.Refresh
Call AddPropertyToFieldDAO( "usuario","idUsuario","Caption","idUsuario",dbText)
Call AddPropertyToFieldDAO( "usuario","nombre","Caption","nombre",dbText)
Call AddPropertyToFieldDAO( "usuario","aPaterno","Caption","aPaterno",dbText)
Call AddPropertyToFieldDAO( "usuario","aMAterno","Caption","aMAterno",dbText)
Call AddPropertyToFieldDAO( "usuario","telefonoMovil","Caption","telefonoMovil",dbText)
Call AddPropertyToFieldDAO( "usuario","telefonoFijo","Caption","telefonoFijo",dbText)
Call AddPropertyToFieldDAO( "usuario","direccion","Caption","direccion",dbText)
Call AddPropertyToFieldDAO( "usuario","fechaIngreso","Caption","fechaIngreso",dbText)
Call AddPropertyToFieldDAO( "usuario","idEstatus","Caption","idEstatus",dbText)

End Sub


'=== Create table nivel ========

Sub CreateTable2()
Set tbl = New ADOX.Table

tbl.Name = "nivel"
cat.Tables.Append tbl

Call AddFieldToTable("nivel", "idNivel", adInteger, 0, 0, "Autoincrement", "", "", "", "", "", "")
Call AddFieldToTable("nivel", "nivel", adSmallInt, 0, 0, "", "TRUE", "", "", "", "", "")
Call AddFieldToTable("nivel", "nombreNivel", adVarWChar, 50, 0, "", "TRUE", "", "", "", "", "")
Call AddFieldToTable("nivel", "descripcion", adVarWChar, 100, 0, "", "TRUE", "", "", "", "", "")


dbs.TableDefs.Refresh
Call AddPropertyToFieldDAO( "nivel","idNivel","Caption","idNivel",dbText)
Call AddPropertyToFieldDAO( "nivel","nivel","Caption","nivel",dbText)
Call AddPropertyToFieldDAO( "nivel","nombreNivel","Caption","nombreNivel",dbText)
Call AddPropertyToFieldDAO( "nivel","descripcion","Caption","descripcion",dbText)

End Sub


'=== Create table color ========

Sub CreateTable3()
Set tbl = New ADOX.Table

tbl.Name = "color"
cat.Tables.Append tbl

Call AddFieldToTable("color", "idColor", adInteger, 0, 0, "Autoincrement", "", "", "", "", "", "")
Call AddFieldToTable("color", "nombre", adVarWChar, 50, 0, "", "TRUE", "", "", "", "", "")
Call AddFieldToTable("color", "activo", adBoolean, 0, 0, "", "", "", "", "", "", "")


dbs.TableDefs.Refresh
Call AddPropertyToFieldDAO( "color","idColor","Caption","idColor",dbText)
Call AddPropertyToFieldDAO( "color","nombre","Caption","nombre",dbText)
Call AddPropertyToFieldDAO( "color","activo","Caption","activo",dbText)

End Sub


'=== Create table nivel_color ========

Sub CreateTable4()
Set tbl = New ADOX.Table

tbl.Name = "nivel_color"
cat.Tables.Append tbl

Call AddFieldToTable("nivel_color", "idNivel", adInteger, 0, 0, "", "FALSE", "", "", "", "", "")
Call AddFieldToTable("nivel_color", "idColor", adInteger, 0, 0, "", "FALSE", "", "", "", "", "")


dbs.TableDefs.Refresh
Call AddPropertyToFieldDAO( "nivel_color","idNivel","Caption","idNivel",dbText)
Call AddPropertyToFieldDAO( "nivel_color","idColor","Caption","idColor",dbText)

End Sub


'=== Create table panel_participacion ========

Sub CreateTable5()
Set tbl = New ADOX.Table

tbl.Name = "panel_participacion"
cat.Tables.Append tbl

Call AddFieldToTable("panel_participacion", "idPanelParticipacion", adInteger, 0, 0, "Autoincrement", "", "", "", "", "", "")
Call AddFieldToTable("panel_participacion", "fechaInicio", adDate, 0, 0, "", "TRUE", "", "", "", "", "")
Call AddFieldToTable("panel_participacion", "idUsuario", adInteger, 0, 0, "", "FALSE", "", "", "", "", "")
Call AddFieldToTable("panel_participacion", "idNivel", adInteger, 0, 0, "", "FALSE", "", "", "", "", "")
Call AddFieldToTable("panel_participacion", "idColor", adInteger, 0, 0, "", "FALSE", "", "", "", "", "")
Call AddFieldToTable("panel_participacion", "idEstatus", adSmallInt, 0, 0, "", "FALSE", "", "", "", "", "")


dbs.TableDefs.Refresh
Call AddPropertyToFieldDAO( "panel_participacion","idPanelParticipacion","Caption","idPanelParticipacion",dbText)
Call AddPropertyToFieldDAO( "panel_participacion","fechaInicio","Caption","fechaInicio",dbText)
Call AddPropertyToFieldDAO( "panel_participacion","idUsuario","Caption","idUsuario",dbText)
Call AddPropertyToFieldDAO( "panel_participacion","idNivel","Caption","idNivel",dbText)
Call AddPropertyToFieldDAO( "panel_participacion","idColor","Caption","idColor",dbText)
Call AddPropertyToFieldDAO( "panel_participacion","idEstatus","Caption","idEstatus",dbText)

End Sub


'=== Create table estatus ========

Sub CreateTable6()
Set tbl = New ADOX.Table

tbl.Name = "estatus"
cat.Tables.Append tbl

Call AddFieldToTable("estatus", "idEstatus", adSmallInt, 0, 0, "", "FALSE", "", "", "", "", "")
Call AddFieldToTable("estatus", "nombre", adVarWChar, 50, 0, "", "TRUE", "", "", "", "", "")
Call AddFieldToTable("estatus", "descripcion", adVarWChar, 200, 0, "", "TRUE", "", "", "", "", "")


dbs.TableDefs.Refresh
Call AddPropertyToFieldDAO( "estatus","idEstatus","Caption","idEstatus",dbText)
Call AddPropertyToFieldDAO( "estatus","nombre","Caption","nombre",dbText)
Call AddPropertyToFieldDAO( "estatus","descripcion","Caption","descripcion",dbText)

End Sub


' Create primary keys
'=====================

Sub CreatePrimaryKeys()

'=== Create primary key for table usuario ======

Set idx = New ADOX.Index
Set tbl = cat.Tables("usuario")

idx.Name = "Key1"
idx.PrimaryKey = True
idx.Unique  = True
idx.IndexNulls = adIndexNullsDisallow

Call AddFieldToIndex(idx, "idUsuario", False )
Call AddFieldToIndex(idx, "idEstatus", False )

tbl.Indexes.Append idx

'=== Create primary key for table nivel ======

Set idx = New ADOX.Index
Set tbl = cat.Tables("nivel")

idx.Name = "Key2"
idx.PrimaryKey = True
idx.Unique  = True
idx.IndexNulls = adIndexNullsDisallow

Call AddFieldToIndex(idx, "idNivel", False )

tbl.Indexes.Append idx

'=== Create primary key for table color ======

Set idx = New ADOX.Index
Set tbl = cat.Tables("color")

idx.Name = "Key3"
idx.PrimaryKey = True
idx.Unique  = True
idx.IndexNulls = adIndexNullsDisallow

Call AddFieldToIndex(idx, "idColor", False )

tbl.Indexes.Append idx

'=== Create primary key for table nivel_color ======

Set idx = New ADOX.Index
Set tbl = cat.Tables("nivel_color")

idx.Name = "Key4"
idx.PrimaryKey = True
idx.Unique  = True
idx.IndexNulls = adIndexNullsDisallow

Call AddFieldToIndex(idx, "idNivel", False )
Call AddFieldToIndex(idx, "idColor", False )

tbl.Indexes.Append idx

'=== Create primary key for table panel_participacion ======

Set idx = New ADOX.Index
Set tbl = cat.Tables("panel_participacion")

idx.Name = "Key5"
idx.PrimaryKey = True
idx.Unique  = True
idx.IndexNulls = adIndexNullsDisallow

Call AddFieldToIndex(idx, "idPanelParticipacion", False )
Call AddFieldToIndex(idx, "idUsuario", False )
Call AddFieldToIndex(idx, "idNivel", False )
Call AddFieldToIndex(idx, "idColor", False )
Call AddFieldToIndex(idx, "idEstatus", False )

tbl.Indexes.Append idx

'=== Create primary key for table estatus ======

Set idx = New ADOX.Index
Set tbl = cat.Tables("estatus")

idx.Name = "Key6"
idx.PrimaryKey = True
idx.Unique  = True
idx.IndexNulls = adIndexNullsDisallow

Call AddFieldToIndex(idx, "idEstatus", False )

tbl.Indexes.Append idx

End Sub

' Create indexes
'================

Sub CreateIndexes()

End Sub

' Create alternate keys (unique indexes in Microsoft Access)
'=====================================================

Sub CreateAlterKeys()

'=== Create alternate key (unique index) for table usuario ======

Set idx = New ADOX.Index
Set tbl = cat.Tables("usuario")

idx.Name = "idUsuario"
idx.Unique  = True
idx.IndexNulls = adIndexNullsAllow

Call AddFieldToIndex(idx, "idUsuario", False )

tbl.Indexes.Append idx

'=== Create alternate key (unique index) for table nivel ======

Set idx = New ADOX.Index
Set tbl = cat.Tables("nivel")

idx.Name = "idNivel"
idx.Unique  = True
idx.IndexNulls = adIndexNullsAllow

Call AddFieldToIndex(idx, "idNivel", False )

tbl.Indexes.Append idx

'=== Create alternate key (unique index) for table color ======

Set idx = New ADOX.Index
Set tbl = cat.Tables("color")

idx.Name = "idColor"
idx.Unique  = True
idx.IndexNulls = adIndexNullsAllow

Call AddFieldToIndex(idx, "idColor", False )

tbl.Indexes.Append idx

'=== Create alternate key (unique index) for table panel_participacion ======

Set idx = New ADOX.Index
Set tbl = cat.Tables("panel_participacion")

idx.Name = "idParticipacion"
idx.Unique  = True
idx.IndexNulls = adIndexNullsAllow

Call AddFieldToIndex(idx, "idPanelParticipacion", False )

tbl.Indexes.Append idx

'=== Create alternate key (unique index) for table estatus ======

Set idx = New ADOX.Index
Set tbl = cat.Tables("estatus")

idx.Name = "idEstatus"
idx.Unique  = True
idx.IndexNulls = adIndexNullsAllow

Call AddFieldToIndex(idx, "idEstatus", False )

tbl.Indexes.Append idx

End Sub

' Create relations
'==================

Sub CreateRelations()

Dim keyFk As ADOX.Key

'=== Create relationship between parent table nivel and child table nivel_color ======

Set keyFk = New ADOX.Key
Set tbl = cat.Tables("nivel_color")

keyFk.Name = "Relationship1"
keyFk.Type = adKeyForeign
keyFk.RelatedTable = "nivel"

Call AddFieldToRelation(keyFk, "idNivel", "idNivel")

tbl.Keys.Append keyFk


'=== Create relationship between parent table color and child table nivel_color ======

Set keyFk = New ADOX.Key
Set tbl = cat.Tables("nivel_color")

keyFk.Name = "Relationship2"
keyFk.Type = adKeyForeign
keyFk.RelatedTable = "color"

Call AddFieldToRelation(keyFk, "idColor", "idColor")

tbl.Keys.Append keyFk


'=== Create relationship between parent table usuario and child table panel_participacion ======

Set keyFk = New ADOX.Key
Set tbl = cat.Tables("panel_participacion")

keyFk.Name = "Relationship3"
keyFk.Type = adKeyForeign
keyFk.RelatedTable = "usuario"

Call AddFieldToRelation(keyFk, "idUsuario", "idUsuario")
Call AddFieldToRelation(keyFk, "idEstatus", "idEstatus")

tbl.Keys.Append keyFk


'=== Create relationship between parent table nivel_color and child table panel_participacion ======

Set keyFk = New ADOX.Key
Set tbl = cat.Tables("panel_participacion")

keyFk.Name = "Relationship4"
keyFk.Type = adKeyForeign
keyFk.RelatedTable = "nivel_color"

Call AddFieldToRelation(keyFk, "idNivel", "idNivel")
Call AddFieldToRelation(keyFk, "idColor", "idColor")

tbl.Keys.Append keyFk


'=== Create relationship between parent table estatus and child table usuario ======

Set keyFk = New ADOX.Key
Set tbl = cat.Tables("usuario")

keyFk.Name = "Relationship6"
keyFk.Type = adKeyForeign
keyFk.RelatedTable = "estatus"

Call AddFieldToRelation(keyFk, "idEstatus", "idEstatus")

tbl.Keys.Append keyFk


End Sub

' Create relations (DAO)
'========================

Sub CreateRelationsDAO()

Dim rel As DAO.Relation
dbs.Relations.Refresh

End Sub

' Create queries
'================

Sub CreateQueries()

Dim cmd As ADODB.Command
End Sub

' Add fields to table
'=====================

Sub AddFieldToTable(TableName As String, FieldName As String, _
              DataType As Integer, SizePrecCol As Integer, ScaleCol As Integer, Attributes As String, _
              Nullable As String, DefaultValue As Variant, _
              AllowZeroLength As String, CompressUnicode As String, _
              ValText As String, ValRule As String)

Dim col As New ADOX.Column

col.ParentCatalog = cat
col.Name = FieldName
col.Type = DataType
If DataType = adVarWChar Then col.DefinedSize = SizePrecCol
If DataType = adNumeric Then
  col.Precision = SizePrecCol
  col.NumericScale = ScaleCol
End If
If Nullable <> "" Then col.Properties("Nullable").Value = CBool(Nullable)
If Attributes <> "" Then col.Properties(Attributes).Value = True
If AllowZeroLength <> "" Then col.Properties("Jet OLEDB:Allow Zero Length").Value = CBool(AllowZeroLength)
If CompressUnicode <> "" Then col.Properties("Jet OLEDB:Compressed UNICODE Strings").Value = CBool(CompressUnicode)
If DefaultValue <> "" Then col.Properties("Default").Value = DefaultValue
If ValRule <> "" Then col.Properties("Jet OLEDB:Column Validation Rule").Value = ValRule
If ValText <> "" Then col.Properties("Jet OLEDB:Column Validation Text").Value = ValText

cat.Tables(TableName).Columns.Append col
End Sub

' Add DAO properties to table
'=============================

Sub AddPropertyToTableDAO( TableName As String, PropertyName As String, Value As Variant, DataType As String)

Dim prp As DAO.Property
Dim tdf As DAO.TableDef

Set tdf = dbs.TableDefs(TableName)
Set prp = tdf.CreateProperty(PropertyName, DataType, Value)

tdf.Properties.Append prp

End Sub

' Add DAO properties to field
'=============================

Sub AddPropertyToFieldDAO( TableName As String, FieldName As String, PropertyName As String, Value As Variant, DataType As String)

Dim prp As DAO.Property
Dim fld As DAO.Field
Dim tdf As DAO.TableDef

Set tdf = dbs.TableDefs(TableName)
Set fld = tdf.Fields( FieldName )
Set prp = fld.CreateProperty(PropertyName, DataType, Value)

fld.Properties.Append prp

End Sub

' Add fields to index
'=====================

Sub AddFieldToIndex(idx As ADOX.Index, FieldName As String, Descending As Boolean )

idx.Columns.Append FieldName
If Descending = True Then idx.Columns(FieldName).SortOrder = adSortDescending

End Sub

' Add fields to relation
'========================

Sub AddFieldToRelation(keyFk As ADOX.Key, PKField As String, FKField As String )

keyFk.Columns.Append FKField
keyFk.Columns(FKField).RelatedColumn = PKField

End Sub

' Add fields to relation (DAO)
'=============================

Sub AddFieldToRelationDAO( rel As DAO.Relation, PKField As String, FKField As String )

Dim fld As DAO.Field

Set fld = rel.CreateField( PKField )

fld.ForeignName = FKField

rel.Fields.Append fld

End Sub



