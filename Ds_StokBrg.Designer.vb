﻿'------------------------------------------------------------------------------
' <auto-generated>
'     This code was generated by a tool.
'     Runtime Version:2.0.50727.5420
'
'     Changes to this file may cause incorrect behavior and will be lost if
'     the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict Off
Option Explicit On



'''<summary>
'''Represents a strongly typed in-memory cache of data.
'''</summary>
<Global.System.CodeDom.Compiler.GeneratedCodeAttribute("System.Data.Design.TypedDataSetGenerator", "2.0.0.0"),  _
 Global.System.Serializable(),  _
 Global.System.ComponentModel.DesignerCategoryAttribute("code"),  _
 Global.System.ComponentModel.ToolboxItem(true),  _
 Global.System.Xml.Serialization.XmlSchemaProviderAttribute("GetTypedDataSetSchema"),  _
 Global.System.Xml.Serialization.XmlRootAttribute("Ds_StokBrg"),  _
 Global.System.ComponentModel.Design.HelpKeywordAttribute("vs.data.DataSet")>  _
Partial Public Class Ds_StokBrg
    Inherits Global.System.Data.DataSet
    
    Private tableDs_StokBrg As Ds_StokBrgDataTable
    
    Private _schemaSerializationMode As Global.System.Data.SchemaSerializationMode = Global.System.Data.SchemaSerializationMode.IncludeSchema
    
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
    Public Sub New()
        MyBase.New
        Me.BeginInit
        Me.InitClass
        Dim schemaChangedHandler As Global.System.ComponentModel.CollectionChangeEventHandler = AddressOf Me.SchemaChanged
        AddHandler MyBase.Tables.CollectionChanged, schemaChangedHandler
        AddHandler MyBase.Relations.CollectionChanged, schemaChangedHandler
        Me.EndInit
    End Sub
    
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
    Protected Sub New(ByVal info As Global.System.Runtime.Serialization.SerializationInfo, ByVal context As Global.System.Runtime.Serialization.StreamingContext)
        MyBase.New(info, context, false)
        If (Me.IsBinarySerialized(info, context) = true) Then
            Me.InitVars(false)
            Dim schemaChangedHandler1 As Global.System.ComponentModel.CollectionChangeEventHandler = AddressOf Me.SchemaChanged
            AddHandler Me.Tables.CollectionChanged, schemaChangedHandler1
            AddHandler Me.Relations.CollectionChanged, schemaChangedHandler1
            Return
        End If
        Dim strSchema As String = CType(info.GetValue("XmlSchema", GetType(String)),String)
        If (Me.DetermineSchemaSerializationMode(info, context) = Global.System.Data.SchemaSerializationMode.IncludeSchema) Then
            Dim ds As Global.System.Data.DataSet = New Global.System.Data.DataSet
            ds.ReadXmlSchema(New Global.System.Xml.XmlTextReader(New Global.System.IO.StringReader(strSchema)))
            If (Not (ds.Tables("Ds_StokBrg")) Is Nothing) Then
                MyBase.Tables.Add(New Ds_StokBrgDataTable(ds.Tables("Ds_StokBrg")))
            End If
            Me.DataSetName = ds.DataSetName
            Me.Prefix = ds.Prefix
            Me.Namespace = ds.Namespace
            Me.Locale = ds.Locale
            Me.CaseSensitive = ds.CaseSensitive
            Me.EnforceConstraints = ds.EnforceConstraints
            Me.Merge(ds, false, Global.System.Data.MissingSchemaAction.Add)
            Me.InitVars
        Else
            Me.ReadXmlSchema(New Global.System.Xml.XmlTextReader(New Global.System.IO.StringReader(strSchema)))
        End If
        Me.GetSerializationData(info, context)
        Dim schemaChangedHandler As Global.System.ComponentModel.CollectionChangeEventHandler = AddressOf Me.SchemaChanged
        AddHandler MyBase.Tables.CollectionChanged, schemaChangedHandler
        AddHandler Me.Relations.CollectionChanged, schemaChangedHandler
    End Sub
    
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.ComponentModel.Browsable(false),  _
     Global.System.ComponentModel.DesignerSerializationVisibility(Global.System.ComponentModel.DesignerSerializationVisibility.Content)>  _
    Public ReadOnly Property _Ds_StokBrg() As Ds_StokBrgDataTable
        Get
            Return Me.tableDs_StokBrg
        End Get
    End Property
    
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.ComponentModel.BrowsableAttribute(true),  _
     Global.System.ComponentModel.DesignerSerializationVisibilityAttribute(Global.System.ComponentModel.DesignerSerializationVisibility.Visible)>  _
    Public Overrides Property SchemaSerializationMode() As Global.System.Data.SchemaSerializationMode
        Get
            Return Me._schemaSerializationMode
        End Get
        Set
            Me._schemaSerializationMode = value
        End Set
    End Property
    
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.ComponentModel.DesignerSerializationVisibilityAttribute(Global.System.ComponentModel.DesignerSerializationVisibility.Hidden)>  _
    Public Shadows ReadOnly Property Tables() As Global.System.Data.DataTableCollection
        Get
            Return MyBase.Tables
        End Get
    End Property
    
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.ComponentModel.DesignerSerializationVisibilityAttribute(Global.System.ComponentModel.DesignerSerializationVisibility.Hidden)>  _
    Public Shadows ReadOnly Property Relations() As Global.System.Data.DataRelationCollection
        Get
            Return MyBase.Relations
        End Get
    End Property
    
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
    Protected Overrides Sub InitializeDerivedDataSet()
        Me.BeginInit
        Me.InitClass
        Me.EndInit
    End Sub
    
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
    Public Overrides Function Clone() As Global.System.Data.DataSet
        Dim cln As Ds_StokBrg = CType(MyBase.Clone,Ds_StokBrg)
        cln.InitVars
        cln.SchemaSerializationMode = Me.SchemaSerializationMode
        Return cln
    End Function
    
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
    Protected Overrides Function ShouldSerializeTables() As Boolean
        Return false
    End Function
    
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
    Protected Overrides Function ShouldSerializeRelations() As Boolean
        Return false
    End Function
    
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
    Protected Overrides Sub ReadXmlSerializable(ByVal reader As Global.System.Xml.XmlReader)
        If (Me.DetermineSchemaSerializationMode(reader) = Global.System.Data.SchemaSerializationMode.IncludeSchema) Then
            Me.Reset
            Dim ds As Global.System.Data.DataSet = New Global.System.Data.DataSet
            ds.ReadXml(reader)
            If (Not (ds.Tables("Ds_StokBrg")) Is Nothing) Then
                MyBase.Tables.Add(New Ds_StokBrgDataTable(ds.Tables("Ds_StokBrg")))
            End If
            Me.DataSetName = ds.DataSetName
            Me.Prefix = ds.Prefix
            Me.Namespace = ds.Namespace
            Me.Locale = ds.Locale
            Me.CaseSensitive = ds.CaseSensitive
            Me.EnforceConstraints = ds.EnforceConstraints
            Me.Merge(ds, false, Global.System.Data.MissingSchemaAction.Add)
            Me.InitVars
        Else
            Me.ReadXml(reader)
            Me.InitVars
        End If
    End Sub
    
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
    Protected Overrides Function GetSchemaSerializable() As Global.System.Xml.Schema.XmlSchema
        Dim stream As Global.System.IO.MemoryStream = New Global.System.IO.MemoryStream
        Me.WriteXmlSchema(New Global.System.Xml.XmlTextWriter(stream, Nothing))
        stream.Position = 0
        Return Global.System.Xml.Schema.XmlSchema.Read(New Global.System.Xml.XmlTextReader(stream), Nothing)
    End Function
    
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
    Friend Overloads Sub InitVars()
        Me.InitVars(true)
    End Sub
    
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
    Friend Overloads Sub InitVars(ByVal initTable As Boolean)
        Me.tableDs_StokBrg = CType(MyBase.Tables("Ds_StokBrg"),Ds_StokBrgDataTable)
        If (initTable = true) Then
            If (Not (Me.tableDs_StokBrg) Is Nothing) Then
                Me.tableDs_StokBrg.InitVars
            End If
        End If
    End Sub
    
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
    Private Sub InitClass()
        Me.DataSetName = "Ds_StokBrg"
        Me.Prefix = ""
        Me.Namespace = "http://tempuri.org/Ds_StokBrg.xsd"
        Me.EnforceConstraints = true
        Me.SchemaSerializationMode = Global.System.Data.SchemaSerializationMode.IncludeSchema
        Me.tableDs_StokBrg = New Ds_StokBrgDataTable
        MyBase.Tables.Add(Me.tableDs_StokBrg)
    End Sub
    
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
    Private Function ShouldSerialize_Ds_StokBrg() As Boolean
        Return false
    End Function
    
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
    Private Sub SchemaChanged(ByVal sender As Object, ByVal e As Global.System.ComponentModel.CollectionChangeEventArgs)
        If (e.Action = Global.System.ComponentModel.CollectionChangeAction.Remove) Then
            Me.InitVars
        End If
    End Sub
    
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
    Public Shared Function GetTypedDataSetSchema(ByVal xs As Global.System.Xml.Schema.XmlSchemaSet) As Global.System.Xml.Schema.XmlSchemaComplexType
        Dim ds As Ds_StokBrg = New Ds_StokBrg
        Dim type As Global.System.Xml.Schema.XmlSchemaComplexType = New Global.System.Xml.Schema.XmlSchemaComplexType
        Dim sequence As Global.System.Xml.Schema.XmlSchemaSequence = New Global.System.Xml.Schema.XmlSchemaSequence
        Dim any As Global.System.Xml.Schema.XmlSchemaAny = New Global.System.Xml.Schema.XmlSchemaAny
        any.Namespace = ds.Namespace
        sequence.Items.Add(any)
        type.Particle = sequence
        Dim dsSchema As Global.System.Xml.Schema.XmlSchema = ds.GetSchemaSerializable
        If xs.Contains(dsSchema.TargetNamespace) Then
            Dim s1 As Global.System.IO.MemoryStream = New Global.System.IO.MemoryStream
            Dim s2 As Global.System.IO.MemoryStream = New Global.System.IO.MemoryStream
            Try 
                Dim schema As Global.System.Xml.Schema.XmlSchema = Nothing
                dsSchema.Write(s1)
                Dim schemas As Global.System.Collections.IEnumerator = xs.Schemas(dsSchema.TargetNamespace).GetEnumerator
                Do While schemas.MoveNext
                    schema = CType(schemas.Current,Global.System.Xml.Schema.XmlSchema)
                    s2.SetLength(0)
                    schema.Write(s2)
                    If (s1.Length = s2.Length) Then
                        s1.Position = 0
                        s2.Position = 0
                        
                        Do While ((s1.Position <> s1.Length)  _
                                    AndAlso (s1.ReadByte = s2.ReadByte))
                            
                            
                        Loop
                        If (s1.Position = s1.Length) Then
                            Return type
                        End If
                    End If
                    
                Loop
            Finally
                If (Not (s1) Is Nothing) Then
                    s1.Close
                End If
                If (Not (s2) Is Nothing) Then
                    s2.Close
                End If
            End Try
        End If
        xs.Add(dsSchema)
        Return type
    End Function
    
    Public Delegate Sub Ds_StokBrgRowChangeEventHandler(ByVal sender As Object, ByVal e As Ds_StokBrgRowChangeEvent)
    
    '''<summary>
    '''Represents the strongly named DataTable class.
    '''</summary>
    <Global.System.CodeDom.Compiler.GeneratedCodeAttribute("System.Data.Design.TypedDataSetGenerator", "2.0.0.0"),  _
     Global.System.Serializable(),  _
     Global.System.Xml.Serialization.XmlSchemaProviderAttribute("GetTypedTableSchema")>  _
    Partial Public Class Ds_StokBrgDataTable
        Inherits Global.System.Data.TypedTableBase(Of Ds_StokBrgRow)
        
        Private columnKd_Brg As Global.System.Data.DataColumn
        
        Private columnJns_Brg As Global.System.Data.DataColumn
        
        Private columnNm_Brg As Global.System.Data.DataColumn
        
        Private columnSpec As Global.System.Data.DataColumn
        
        Private columnSatuan As Global.System.Data.DataColumn
        
        Private columnSaldo_Akhir As Global.System.Data.DataColumn
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public Sub New()
            MyBase.New
            Me.TableName = "Ds_StokBrg"
            Me.BeginInit
            Me.InitClass
            Me.EndInit
        End Sub
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Friend Sub New(ByVal table As Global.System.Data.DataTable)
            MyBase.New
            Me.TableName = table.TableName
            If (table.CaseSensitive <> table.DataSet.CaseSensitive) Then
                Me.CaseSensitive = table.CaseSensitive
            End If
            If (table.Locale.ToString <> table.DataSet.Locale.ToString) Then
                Me.Locale = table.Locale
            End If
            If (table.Namespace <> table.DataSet.Namespace) Then
                Me.Namespace = table.Namespace
            End If
            Me.Prefix = table.Prefix
            Me.MinimumCapacity = table.MinimumCapacity
        End Sub
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Protected Sub New(ByVal info As Global.System.Runtime.Serialization.SerializationInfo, ByVal context As Global.System.Runtime.Serialization.StreamingContext)
            MyBase.New(info, context)
            Me.InitVars
        End Sub
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public ReadOnly Property Kd_BrgColumn() As Global.System.Data.DataColumn
            Get
                Return Me.columnKd_Brg
            End Get
        End Property
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public ReadOnly Property Jns_BrgColumn() As Global.System.Data.DataColumn
            Get
                Return Me.columnJns_Brg
            End Get
        End Property
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public ReadOnly Property Nm_BrgColumn() As Global.System.Data.DataColumn
            Get
                Return Me.columnNm_Brg
            End Get
        End Property
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public ReadOnly Property SpecColumn() As Global.System.Data.DataColumn
            Get
                Return Me.columnSpec
            End Get
        End Property
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public ReadOnly Property SatuanColumn() As Global.System.Data.DataColumn
            Get
                Return Me.columnSatuan
            End Get
        End Property
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public ReadOnly Property Saldo_AkhirColumn() As Global.System.Data.DataColumn
            Get
                Return Me.columnSaldo_Akhir
            End Get
        End Property
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.ComponentModel.Browsable(false)>  _
        Public ReadOnly Property Count() As Integer
            Get
                Return Me.Rows.Count
            End Get
        End Property
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public Default ReadOnly Property Item(ByVal index As Integer) As Ds_StokBrgRow
            Get
                Return CType(Me.Rows(index),Ds_StokBrgRow)
            End Get
        End Property
        
        Public Event Ds_StokBrgRowChanging As Ds_StokBrgRowChangeEventHandler
        
        Public Event Ds_StokBrgRowChanged As Ds_StokBrgRowChangeEventHandler
        
        Public Event Ds_StokBrgRowDeleting As Ds_StokBrgRowChangeEventHandler
        
        Public Event Ds_StokBrgRowDeleted As Ds_StokBrgRowChangeEventHandler
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public Overloads Sub AddDs_StokBrgRow(ByVal row As Ds_StokBrgRow)
            Me.Rows.Add(row)
        End Sub
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public Overloads Function AddDs_StokBrgRow(ByVal Kd_Brg As String, ByVal Jns_Brg As String, ByVal Nm_Brg As String, ByVal Spec As String, ByVal Satuan As String, ByVal Saldo_Akhir As String) As Ds_StokBrgRow
            Dim rowDs_StokBrgRow As Ds_StokBrgRow = CType(Me.NewRow,Ds_StokBrgRow)
            Dim columnValuesArray() As Object = New Object() {Kd_Brg, Jns_Brg, Nm_Brg, Spec, Satuan, Saldo_Akhir}
            rowDs_StokBrgRow.ItemArray = columnValuesArray
            Me.Rows.Add(rowDs_StokBrgRow)
            Return rowDs_StokBrgRow
        End Function
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public Overrides Function Clone() As Global.System.Data.DataTable
            Dim cln As Ds_StokBrgDataTable = CType(MyBase.Clone,Ds_StokBrgDataTable)
            cln.InitVars
            Return cln
        End Function
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Protected Overrides Function CreateInstance() As Global.System.Data.DataTable
            Return New Ds_StokBrgDataTable
        End Function
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Friend Sub InitVars()
            Me.columnKd_Brg = MyBase.Columns("Kd_Brg")
            Me.columnJns_Brg = MyBase.Columns("Jns_Brg")
            Me.columnNm_Brg = MyBase.Columns("Nm_Brg")
            Me.columnSpec = MyBase.Columns("Spec")
            Me.columnSatuan = MyBase.Columns("Satuan")
            Me.columnSaldo_Akhir = MyBase.Columns("Saldo_Akhir")
        End Sub
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Private Sub InitClass()
            Me.columnKd_Brg = New Global.System.Data.DataColumn("Kd_Brg", GetType(String), Nothing, Global.System.Data.MappingType.Element)
            MyBase.Columns.Add(Me.columnKd_Brg)
            Me.columnJns_Brg = New Global.System.Data.DataColumn("Jns_Brg", GetType(String), Nothing, Global.System.Data.MappingType.Element)
            MyBase.Columns.Add(Me.columnJns_Brg)
            Me.columnNm_Brg = New Global.System.Data.DataColumn("Nm_Brg", GetType(String), Nothing, Global.System.Data.MappingType.Element)
            MyBase.Columns.Add(Me.columnNm_Brg)
            Me.columnSpec = New Global.System.Data.DataColumn("Spec", GetType(String), Nothing, Global.System.Data.MappingType.Element)
            MyBase.Columns.Add(Me.columnSpec)
            Me.columnSatuan = New Global.System.Data.DataColumn("Satuan", GetType(String), Nothing, Global.System.Data.MappingType.Element)
            MyBase.Columns.Add(Me.columnSatuan)
            Me.columnSaldo_Akhir = New Global.System.Data.DataColumn("Saldo_Akhir", GetType(String), Nothing, Global.System.Data.MappingType.Element)
            MyBase.Columns.Add(Me.columnSaldo_Akhir)
            Me.ExtendedProperties.Add("Generator_TablePropName", "_Ds_StokBrg")
            Me.ExtendedProperties.Add("Generator_UserTableName", "Ds_StokBrg")
        End Sub
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public Function NewDs_StokBrgRow() As Ds_StokBrgRow
            Return CType(Me.NewRow,Ds_StokBrgRow)
        End Function
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Protected Overrides Function NewRowFromBuilder(ByVal builder As Global.System.Data.DataRowBuilder) As Global.System.Data.DataRow
            Return New Ds_StokBrgRow(builder)
        End Function
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Protected Overrides Function GetRowType() As Global.System.Type
            Return GetType(Ds_StokBrgRow)
        End Function
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Protected Overrides Sub OnRowChanged(ByVal e As Global.System.Data.DataRowChangeEventArgs)
            MyBase.OnRowChanged(e)
            If (Not (Me.Ds_StokBrgRowChangedEvent) Is Nothing) Then
                RaiseEvent Ds_StokBrgRowChanged(Me, New Ds_StokBrgRowChangeEvent(CType(e.Row,Ds_StokBrgRow), e.Action))
            End If
        End Sub
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Protected Overrides Sub OnRowChanging(ByVal e As Global.System.Data.DataRowChangeEventArgs)
            MyBase.OnRowChanging(e)
            If (Not (Me.Ds_StokBrgRowChangingEvent) Is Nothing) Then
                RaiseEvent Ds_StokBrgRowChanging(Me, New Ds_StokBrgRowChangeEvent(CType(e.Row,Ds_StokBrgRow), e.Action))
            End If
        End Sub
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Protected Overrides Sub OnRowDeleted(ByVal e As Global.System.Data.DataRowChangeEventArgs)
            MyBase.OnRowDeleted(e)
            If (Not (Me.Ds_StokBrgRowDeletedEvent) Is Nothing) Then
                RaiseEvent Ds_StokBrgRowDeleted(Me, New Ds_StokBrgRowChangeEvent(CType(e.Row,Ds_StokBrgRow), e.Action))
            End If
        End Sub
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Protected Overrides Sub OnRowDeleting(ByVal e As Global.System.Data.DataRowChangeEventArgs)
            MyBase.OnRowDeleting(e)
            If (Not (Me.Ds_StokBrgRowDeletingEvent) Is Nothing) Then
                RaiseEvent Ds_StokBrgRowDeleting(Me, New Ds_StokBrgRowChangeEvent(CType(e.Row,Ds_StokBrgRow), e.Action))
            End If
        End Sub
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public Sub RemoveDs_StokBrgRow(ByVal row As Ds_StokBrgRow)
            Me.Rows.Remove(row)
        End Sub
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public Shared Function GetTypedTableSchema(ByVal xs As Global.System.Xml.Schema.XmlSchemaSet) As Global.System.Xml.Schema.XmlSchemaComplexType
            Dim type As Global.System.Xml.Schema.XmlSchemaComplexType = New Global.System.Xml.Schema.XmlSchemaComplexType
            Dim sequence As Global.System.Xml.Schema.XmlSchemaSequence = New Global.System.Xml.Schema.XmlSchemaSequence
            Dim ds As Ds_StokBrg = New Ds_StokBrg
            Dim any1 As Global.System.Xml.Schema.XmlSchemaAny = New Global.System.Xml.Schema.XmlSchemaAny
            any1.Namespace = "http://www.w3.org/2001/XMLSchema"
            any1.MinOccurs = New Decimal(0)
            any1.MaxOccurs = Decimal.MaxValue
            any1.ProcessContents = Global.System.Xml.Schema.XmlSchemaContentProcessing.Lax
            sequence.Items.Add(any1)
            Dim any2 As Global.System.Xml.Schema.XmlSchemaAny = New Global.System.Xml.Schema.XmlSchemaAny
            any2.Namespace = "urn:schemas-microsoft-com:xml-diffgram-v1"
            any2.MinOccurs = New Decimal(1)
            any2.ProcessContents = Global.System.Xml.Schema.XmlSchemaContentProcessing.Lax
            sequence.Items.Add(any2)
            Dim attribute1 As Global.System.Xml.Schema.XmlSchemaAttribute = New Global.System.Xml.Schema.XmlSchemaAttribute
            attribute1.Name = "namespace"
            attribute1.FixedValue = ds.Namespace
            type.Attributes.Add(attribute1)
            Dim attribute2 As Global.System.Xml.Schema.XmlSchemaAttribute = New Global.System.Xml.Schema.XmlSchemaAttribute
            attribute2.Name = "tableTypeName"
            attribute2.FixedValue = "Ds_StokBrgDataTable"
            type.Attributes.Add(attribute2)
            type.Particle = sequence
            Dim dsSchema As Global.System.Xml.Schema.XmlSchema = ds.GetSchemaSerializable
            If xs.Contains(dsSchema.TargetNamespace) Then
                Dim s1 As Global.System.IO.MemoryStream = New Global.System.IO.MemoryStream
                Dim s2 As Global.System.IO.MemoryStream = New Global.System.IO.MemoryStream
                Try 
                    Dim schema As Global.System.Xml.Schema.XmlSchema = Nothing
                    dsSchema.Write(s1)
                    Dim schemas As Global.System.Collections.IEnumerator = xs.Schemas(dsSchema.TargetNamespace).GetEnumerator
                    Do While schemas.MoveNext
                        schema = CType(schemas.Current,Global.System.Xml.Schema.XmlSchema)
                        s2.SetLength(0)
                        schema.Write(s2)
                        If (s1.Length = s2.Length) Then
                            s1.Position = 0
                            s2.Position = 0
                            
                            Do While ((s1.Position <> s1.Length)  _
                                        AndAlso (s1.ReadByte = s2.ReadByte))
                                
                                
                            Loop
                            If (s1.Position = s1.Length) Then
                                Return type
                            End If
                        End If
                        
                    Loop
                Finally
                    If (Not (s1) Is Nothing) Then
                        s1.Close
                    End If
                    If (Not (s2) Is Nothing) Then
                        s2.Close
                    End If
                End Try
            End If
            xs.Add(dsSchema)
            Return type
        End Function
    End Class
    
    '''<summary>
    '''Represents strongly named DataRow class.
    '''</summary>
    <Global.System.CodeDom.Compiler.GeneratedCodeAttribute("System.Data.Design.TypedDataSetGenerator", "2.0.0.0")>  _
    Partial Public Class Ds_StokBrgRow
        Inherits Global.System.Data.DataRow
        
        Private tableDs_StokBrg As Ds_StokBrgDataTable
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Friend Sub New(ByVal rb As Global.System.Data.DataRowBuilder)
            MyBase.New(rb)
            Me.tableDs_StokBrg = CType(Me.Table,Ds_StokBrgDataTable)
        End Sub
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public Property Kd_Brg() As String
            Get
                Try 
                    Return CType(Me(Me.tableDs_StokBrg.Kd_BrgColumn),String)
                Catch e As Global.System.InvalidCastException
                    Throw New Global.System.Data.StrongTypingException("The value for column 'Kd_Brg' in table 'Ds_StokBrg' is DBNull.", e)
                End Try
            End Get
            Set
                Me(Me.tableDs_StokBrg.Kd_BrgColumn) = value
            End Set
        End Property
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public Property Jns_Brg() As String
            Get
                Try 
                    Return CType(Me(Me.tableDs_StokBrg.Jns_BrgColumn),String)
                Catch e As Global.System.InvalidCastException
                    Throw New Global.System.Data.StrongTypingException("The value for column 'Jns_Brg' in table 'Ds_StokBrg' is DBNull.", e)
                End Try
            End Get
            Set
                Me(Me.tableDs_StokBrg.Jns_BrgColumn) = value
            End Set
        End Property
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public Property Nm_Brg() As String
            Get
                Try 
                    Return CType(Me(Me.tableDs_StokBrg.Nm_BrgColumn),String)
                Catch e As Global.System.InvalidCastException
                    Throw New Global.System.Data.StrongTypingException("The value for column 'Nm_Brg' in table 'Ds_StokBrg' is DBNull.", e)
                End Try
            End Get
            Set
                Me(Me.tableDs_StokBrg.Nm_BrgColumn) = value
            End Set
        End Property
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public Property Spec() As String
            Get
                Try 
                    Return CType(Me(Me.tableDs_StokBrg.SpecColumn),String)
                Catch e As Global.System.InvalidCastException
                    Throw New Global.System.Data.StrongTypingException("The value for column 'Spec' in table 'Ds_StokBrg' is DBNull.", e)
                End Try
            End Get
            Set
                Me(Me.tableDs_StokBrg.SpecColumn) = value
            End Set
        End Property
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public Property Satuan() As String
            Get
                Try 
                    Return CType(Me(Me.tableDs_StokBrg.SatuanColumn),String)
                Catch e As Global.System.InvalidCastException
                    Throw New Global.System.Data.StrongTypingException("The value for column 'Satuan' in table 'Ds_StokBrg' is DBNull.", e)
                End Try
            End Get
            Set
                Me(Me.tableDs_StokBrg.SatuanColumn) = value
            End Set
        End Property
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public Property Saldo_Akhir() As String
            Get
                Try 
                    Return CType(Me(Me.tableDs_StokBrg.Saldo_AkhirColumn),String)
                Catch e As Global.System.InvalidCastException
                    Throw New Global.System.Data.StrongTypingException("The value for column 'Saldo_Akhir' in table 'Ds_StokBrg' is DBNull.", e)
                End Try
            End Get
            Set
                Me(Me.tableDs_StokBrg.Saldo_AkhirColumn) = value
            End Set
        End Property
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public Function IsKd_BrgNull() As Boolean
            Return Me.IsNull(Me.tableDs_StokBrg.Kd_BrgColumn)
        End Function
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public Sub SetKd_BrgNull()
            Me(Me.tableDs_StokBrg.Kd_BrgColumn) = Global.System.Convert.DBNull
        End Sub
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public Function IsJns_BrgNull() As Boolean
            Return Me.IsNull(Me.tableDs_StokBrg.Jns_BrgColumn)
        End Function
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public Sub SetJns_BrgNull()
            Me(Me.tableDs_StokBrg.Jns_BrgColumn) = Global.System.Convert.DBNull
        End Sub
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public Function IsNm_BrgNull() As Boolean
            Return Me.IsNull(Me.tableDs_StokBrg.Nm_BrgColumn)
        End Function
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public Sub SetNm_BrgNull()
            Me(Me.tableDs_StokBrg.Nm_BrgColumn) = Global.System.Convert.DBNull
        End Sub
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public Function IsSpecNull() As Boolean
            Return Me.IsNull(Me.tableDs_StokBrg.SpecColumn)
        End Function
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public Sub SetSpecNull()
            Me(Me.tableDs_StokBrg.SpecColumn) = Global.System.Convert.DBNull
        End Sub
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public Function IsSatuanNull() As Boolean
            Return Me.IsNull(Me.tableDs_StokBrg.SatuanColumn)
        End Function
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public Sub SetSatuanNull()
            Me(Me.tableDs_StokBrg.SatuanColumn) = Global.System.Convert.DBNull
        End Sub
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public Function IsSaldo_AkhirNull() As Boolean
            Return Me.IsNull(Me.tableDs_StokBrg.Saldo_AkhirColumn)
        End Function
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public Sub SetSaldo_AkhirNull()
            Me(Me.tableDs_StokBrg.Saldo_AkhirColumn) = Global.System.Convert.DBNull
        End Sub
    End Class
    
    '''<summary>
    '''Row event argument class
    '''</summary>
    <Global.System.CodeDom.Compiler.GeneratedCodeAttribute("System.Data.Design.TypedDataSetGenerator", "2.0.0.0")>  _
    Public Class Ds_StokBrgRowChangeEvent
        Inherits Global.System.EventArgs
        
        Private eventRow As Ds_StokBrgRow
        
        Private eventAction As Global.System.Data.DataRowAction
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public Sub New(ByVal row As Ds_StokBrgRow, ByVal action As Global.System.Data.DataRowAction)
            MyBase.New
            Me.eventRow = row
            Me.eventAction = action
        End Sub
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public ReadOnly Property Row() As Ds_StokBrgRow
            Get
                Return Me.eventRow
            End Get
        End Property
        
        <Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public ReadOnly Property Action() As Global.System.Data.DataRowAction
            Get
                Return Me.eventAction
            End Get
        End Property
    End Class
End Class
