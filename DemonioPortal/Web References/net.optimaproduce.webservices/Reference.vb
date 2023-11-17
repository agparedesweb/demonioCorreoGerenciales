﻿'------------------------------------------------------------------------------
' <auto-generated>
'     Este código fue generado por una herramienta.
'     Versión de runtime:4.0.30319.42000
'
'     Los cambios en este archivo podrían causar un comportamiento incorrecto y se perderán si
'     se vuelve a generar el código.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict Off
Option Explicit On

Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Diagnostics
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Xml.Serialization

'
'Microsoft.VSDesigner generó automáticamente este código fuente, versión=4.0.30319.42000.
'
Namespace net.optimaproduce.webservices
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.8.9032.0"),  _
     System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code"),  _
     System.Web.Services.WebServiceBindingAttribute(Name:="ServiceSoap", [Namespace]:="http://tempuri.org/")>  _
    Partial Public Class Service
        Inherits System.Web.Services.Protocols.SoapHttpClientProtocol
        
        Private GetSalesOperationCompleted As System.Threading.SendOrPostCallback
        
        Private GetAdjustementsOperationCompleted As System.Threading.SendOrPostCallback
        
        Private GetInventoryActivityOperationCompleted As System.Threading.SendOrPostCallback
        
        Private GetTodayInvoicesOperationCompleted As System.Threading.SendOrPostCallback
        
        Private GetGeneralExpensesOperationCompleted As System.Threading.SendOrPostCallback
        
        Private GetLotExpensesOperationCompleted As System.Threading.SendOrPostCallback
        
        Private useDefaultCredentialsSetExplicitly As Boolean
        
        '''<remarks/>
        Public Sub New()
            MyBase.New
            Me.Url = Global.IntegraFacturacion.My.MySettings.Default.DemonioPortal_net_optimaproduce_webservices_Service
            If (Me.IsLocalFileSystemWebService(Me.Url) = true) Then
                Me.UseDefaultCredentials = true
                Me.useDefaultCredentialsSetExplicitly = false
            Else
                Me.useDefaultCredentialsSetExplicitly = true
            End If
        End Sub
        
        Public Shadows Property Url() As String
            Get
                Return MyBase.Url
            End Get
            Set
                If (((Me.IsLocalFileSystemWebService(MyBase.Url) = true)  _
                            AndAlso (Me.useDefaultCredentialsSetExplicitly = false))  _
                            AndAlso (Me.IsLocalFileSystemWebService(value) = false)) Then
                    MyBase.UseDefaultCredentials = false
                End If
                MyBase.Url = value
            End Set
        End Property
        
        Public Shadows Property UseDefaultCredentials() As Boolean
            Get
                Return MyBase.UseDefaultCredentials
            End Get
            Set
                MyBase.UseDefaultCredentials = value
                Me.useDefaultCredentialsSetExplicitly = true
            End Set
        End Property
        
        '''<remarks/>
        Public Event GetSalesCompleted As GetSalesCompletedEventHandler
        
        '''<remarks/>
        Public Event GetAdjustementsCompleted As GetAdjustementsCompletedEventHandler
        
        '''<remarks/>
        Public Event GetInventoryActivityCompleted As GetInventoryActivityCompletedEventHandler
        
        '''<remarks/>
        Public Event GetTodayInvoicesCompleted As GetTodayInvoicesCompletedEventHandler
        
        '''<remarks/>
        Public Event GetGeneralExpensesCompleted As GetGeneralExpensesCompletedEventHandler
        
        '''<remarks/>
        Public Event GetLotExpensesCompleted As GetLotExpensesCompletedEventHandler
        
        '''<remarks/>
        <System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/GetSales", RequestNamespace:="http://tempuri.org/", ResponseNamespace:="http://tempuri.org/", Use:=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle:=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)>  _
        Public Function GetSales(ByVal OptimaCompanyID As Integer, ByVal companyID As Integer, ByVal yearID As Integer, ByVal batchID As Integer, ByVal packerID As Integer, ByVal growerID As Integer) As System.Data.DataTable
            Dim results() As Object = Me.Invoke("GetSales", New Object() {OptimaCompanyID, companyID, yearID, batchID, packerID, growerID})
            Return CType(results(0),System.Data.DataTable)
        End Function
        
        '''<remarks/>
        Public Overloads Sub GetSalesAsync(ByVal OptimaCompanyID As Integer, ByVal companyID As Integer, ByVal yearID As Integer, ByVal batchID As Integer, ByVal packerID As Integer, ByVal growerID As Integer)
            Me.GetSalesAsync(OptimaCompanyID, companyID, yearID, batchID, packerID, growerID, Nothing)
        End Sub
        
        '''<remarks/>
        Public Overloads Sub GetSalesAsync(ByVal OptimaCompanyID As Integer, ByVal companyID As Integer, ByVal yearID As Integer, ByVal batchID As Integer, ByVal packerID As Integer, ByVal growerID As Integer, ByVal userState As Object)
            If (Me.GetSalesOperationCompleted Is Nothing) Then
                Me.GetSalesOperationCompleted = AddressOf Me.OnGetSalesOperationCompleted
            End If
            Me.InvokeAsync("GetSales", New Object() {OptimaCompanyID, companyID, yearID, batchID, packerID, growerID}, Me.GetSalesOperationCompleted, userState)
        End Sub
        
        Private Sub OnGetSalesOperationCompleted(ByVal arg As Object)
            If (Not (Me.GetSalesCompletedEvent) Is Nothing) Then
                Dim invokeArgs As System.Web.Services.Protocols.InvokeCompletedEventArgs = CType(arg,System.Web.Services.Protocols.InvokeCompletedEventArgs)
                RaiseEvent GetSalesCompleted(Me, New GetSalesCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState))
            End If
        End Sub
        
        '''<remarks/>
        <System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/GetAdjustements", RequestNamespace:="http://tempuri.org/", ResponseNamespace:="http://tempuri.org/", Use:=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle:=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)>  _
        Public Function GetAdjustements(ByVal OptimaCompanyID As Integer, ByVal companyID As Integer, ByVal packerID As Integer, ByVal growerID As Integer, ByVal batchID As Integer) As System.Data.DataTable
            Dim results() As Object = Me.Invoke("GetAdjustements", New Object() {OptimaCompanyID, companyID, packerID, growerID, batchID})
            Return CType(results(0),System.Data.DataTable)
        End Function
        
        '''<remarks/>
        Public Overloads Sub GetAdjustementsAsync(ByVal OptimaCompanyID As Integer, ByVal companyID As Integer, ByVal packerID As Integer, ByVal growerID As Integer, ByVal batchID As Integer)
            Me.GetAdjustementsAsync(OptimaCompanyID, companyID, packerID, growerID, batchID, Nothing)
        End Sub
        
        '''<remarks/>
        Public Overloads Sub GetAdjustementsAsync(ByVal OptimaCompanyID As Integer, ByVal companyID As Integer, ByVal packerID As Integer, ByVal growerID As Integer, ByVal batchID As Integer, ByVal userState As Object)
            If (Me.GetAdjustementsOperationCompleted Is Nothing) Then
                Me.GetAdjustementsOperationCompleted = AddressOf Me.OnGetAdjustementsOperationCompleted
            End If
            Me.InvokeAsync("GetAdjustements", New Object() {OptimaCompanyID, companyID, packerID, growerID, batchID}, Me.GetAdjustementsOperationCompleted, userState)
        End Sub
        
        Private Sub OnGetAdjustementsOperationCompleted(ByVal arg As Object)
            If (Not (Me.GetAdjustementsCompletedEvent) Is Nothing) Then
                Dim invokeArgs As System.Web.Services.Protocols.InvokeCompletedEventArgs = CType(arg,System.Web.Services.Protocols.InvokeCompletedEventArgs)
                RaiseEvent GetAdjustementsCompleted(Me, New GetAdjustementsCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState))
            End If
        End Sub
        
        '''<remarks/>
        <System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/GetInventoryActivity", RequestNamespace:="http://tempuri.org/", ResponseNamespace:="http://tempuri.org/", Use:=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle:=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)>  _
        Public Function GetInventoryActivity(ByVal OptimaCompanyID As Integer, ByVal companyID As Integer, ByVal LdDate As Date, ByVal packerID As Integer, ByVal growerID As Integer, ByVal LabelID As Integer, ByVal Language As String) As System.Data.DataTable
            Dim results() As Object = Me.Invoke("GetInventoryActivity", New Object() {OptimaCompanyID, companyID, LdDate, packerID, growerID, LabelID, Language})
            Return CType(results(0),System.Data.DataTable)
        End Function
        
        '''<remarks/>
        Public Overloads Sub GetInventoryActivityAsync(ByVal OptimaCompanyID As Integer, ByVal companyID As Integer, ByVal LdDate As Date, ByVal packerID As Integer, ByVal growerID As Integer, ByVal LabelID As Integer, ByVal Language As String)
            Me.GetInventoryActivityAsync(OptimaCompanyID, companyID, LdDate, packerID, growerID, LabelID, Language, Nothing)
        End Sub
        
        '''<remarks/>
        Public Overloads Sub GetInventoryActivityAsync(ByVal OptimaCompanyID As Integer, ByVal companyID As Integer, ByVal LdDate As Date, ByVal packerID As Integer, ByVal growerID As Integer, ByVal LabelID As Integer, ByVal Language As String, ByVal userState As Object)
            If (Me.GetInventoryActivityOperationCompleted Is Nothing) Then
                Me.GetInventoryActivityOperationCompleted = AddressOf Me.OnGetInventoryActivityOperationCompleted
            End If
            Me.InvokeAsync("GetInventoryActivity", New Object() {OptimaCompanyID, companyID, LdDate, packerID, growerID, LabelID, Language}, Me.GetInventoryActivityOperationCompleted, userState)
        End Sub
        
        Private Sub OnGetInventoryActivityOperationCompleted(ByVal arg As Object)
            If (Not (Me.GetInventoryActivityCompletedEvent) Is Nothing) Then
                Dim invokeArgs As System.Web.Services.Protocols.InvokeCompletedEventArgs = CType(arg,System.Web.Services.Protocols.InvokeCompletedEventArgs)
                RaiseEvent GetInventoryActivityCompleted(Me, New GetInventoryActivityCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState))
            End If
        End Sub
        
        '''<remarks/>
        <System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/GetTodayInvoices", RequestNamespace:="http://tempuri.org/", ResponseNamespace:="http://tempuri.org/", Use:=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle:=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)>  _
        Public Function GetTodayInvoices(ByVal OptimaCompanyID As Integer, ByVal companyID As Integer, ByVal yearID As Integer, ByVal packerID As Integer, ByVal growerID As Integer, ByVal InvoiceDate As Date, ByVal ContractFilter As Integer) As System.Data.DataTable
            Dim results() As Object = Me.Invoke("GetTodayInvoices", New Object() {OptimaCompanyID, companyID, yearID, packerID, growerID, InvoiceDate, ContractFilter})
            Return CType(results(0),System.Data.DataTable)
        End Function
        
        '''<remarks/>
        Public Overloads Sub GetTodayInvoicesAsync(ByVal OptimaCompanyID As Integer, ByVal companyID As Integer, ByVal yearID As Integer, ByVal packerID As Integer, ByVal growerID As Integer, ByVal InvoiceDate As Date, ByVal ContractFilter As Integer)
            Me.GetTodayInvoicesAsync(OptimaCompanyID, companyID, yearID, packerID, growerID, InvoiceDate, ContractFilter, Nothing)
        End Sub
        
        '''<remarks/>
        Public Overloads Sub GetTodayInvoicesAsync(ByVal OptimaCompanyID As Integer, ByVal companyID As Integer, ByVal yearID As Integer, ByVal packerID As Integer, ByVal growerID As Integer, ByVal InvoiceDate As Date, ByVal ContractFilter As Integer, ByVal userState As Object)
            If (Me.GetTodayInvoicesOperationCompleted Is Nothing) Then
                Me.GetTodayInvoicesOperationCompleted = AddressOf Me.OnGetTodayInvoicesOperationCompleted
            End If
            Me.InvokeAsync("GetTodayInvoices", New Object() {OptimaCompanyID, companyID, yearID, packerID, growerID, InvoiceDate, ContractFilter}, Me.GetTodayInvoicesOperationCompleted, userState)
        End Sub
        
        Private Sub OnGetTodayInvoicesOperationCompleted(ByVal arg As Object)
            If (Not (Me.GetTodayInvoicesCompletedEvent) Is Nothing) Then
                Dim invokeArgs As System.Web.Services.Protocols.InvokeCompletedEventArgs = CType(arg,System.Web.Services.Protocols.InvokeCompletedEventArgs)
                RaiseEvent GetTodayInvoicesCompleted(Me, New GetTodayInvoicesCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState))
            End If
        End Sub
        
        '''<remarks/>
        <System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/GetGeneralExpenses", RequestNamespace:="http://tempuri.org/", ResponseNamespace:="http://tempuri.org/", Use:=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle:=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)>  _
        Public Function GetGeneralExpenses(ByVal OptimaCompanyID As Integer, ByVal companyID As Integer, ByVal yearID As Integer, ByVal packerID As Integer, ByVal growerID As Integer, ByVal batchID As Integer, ByVal expenseID As Integer) As System.Data.DataTable
            Dim results() As Object = Me.Invoke("GetGeneralExpenses", New Object() {OptimaCompanyID, companyID, yearID, packerID, growerID, batchID, expenseID})
            Return CType(results(0),System.Data.DataTable)
        End Function
        
        '''<remarks/>
        Public Overloads Sub GetGeneralExpensesAsync(ByVal OptimaCompanyID As Integer, ByVal companyID As Integer, ByVal yearID As Integer, ByVal packerID As Integer, ByVal growerID As Integer, ByVal batchID As Integer, ByVal expenseID As Integer)
            Me.GetGeneralExpensesAsync(OptimaCompanyID, companyID, yearID, packerID, growerID, batchID, expenseID, Nothing)
        End Sub
        
        '''<remarks/>
        Public Overloads Sub GetGeneralExpensesAsync(ByVal OptimaCompanyID As Integer, ByVal companyID As Integer, ByVal yearID As Integer, ByVal packerID As Integer, ByVal growerID As Integer, ByVal batchID As Integer, ByVal expenseID As Integer, ByVal userState As Object)
            If (Me.GetGeneralExpensesOperationCompleted Is Nothing) Then
                Me.GetGeneralExpensesOperationCompleted = AddressOf Me.OnGetGeneralExpensesOperationCompleted
            End If
            Me.InvokeAsync("GetGeneralExpenses", New Object() {OptimaCompanyID, companyID, yearID, packerID, growerID, batchID, expenseID}, Me.GetGeneralExpensesOperationCompleted, userState)
        End Sub
        
        Private Sub OnGetGeneralExpensesOperationCompleted(ByVal arg As Object)
            If (Not (Me.GetGeneralExpensesCompletedEvent) Is Nothing) Then
                Dim invokeArgs As System.Web.Services.Protocols.InvokeCompletedEventArgs = CType(arg,System.Web.Services.Protocols.InvokeCompletedEventArgs)
                RaiseEvent GetGeneralExpensesCompleted(Me, New GetGeneralExpensesCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState))
            End If
        End Sub
        
        '''<remarks/>
        <System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/GetLotExpenses", RequestNamespace:="http://tempuri.org/", ResponseNamespace:="http://tempuri.org/", Use:=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle:=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)>  _
        Public Function GetLotExpenses(ByVal OptimaCompanyID As Integer, ByVal companyID As Integer, ByVal yearID As Integer, ByVal packerID As Integer, ByVal growerID As Integer, ByVal batchID As Integer, ByVal expenseID As Integer) As System.Data.DataTable
            Dim results() As Object = Me.Invoke("GetLotExpenses", New Object() {OptimaCompanyID, companyID, yearID, packerID, growerID, batchID, expenseID})
            Return CType(results(0),System.Data.DataTable)
        End Function
        
        '''<remarks/>
        Public Overloads Sub GetLotExpensesAsync(ByVal OptimaCompanyID As Integer, ByVal companyID As Integer, ByVal yearID As Integer, ByVal packerID As Integer, ByVal growerID As Integer, ByVal batchID As Integer, ByVal expenseID As Integer)
            Me.GetLotExpensesAsync(OptimaCompanyID, companyID, yearID, packerID, growerID, batchID, expenseID, Nothing)
        End Sub
        
        '''<remarks/>
        Public Overloads Sub GetLotExpensesAsync(ByVal OptimaCompanyID As Integer, ByVal companyID As Integer, ByVal yearID As Integer, ByVal packerID As Integer, ByVal growerID As Integer, ByVal batchID As Integer, ByVal expenseID As Integer, ByVal userState As Object)
            If (Me.GetLotExpensesOperationCompleted Is Nothing) Then
                Me.GetLotExpensesOperationCompleted = AddressOf Me.OnGetLotExpensesOperationCompleted
            End If
            Me.InvokeAsync("GetLotExpenses", New Object() {OptimaCompanyID, companyID, yearID, packerID, growerID, batchID, expenseID}, Me.GetLotExpensesOperationCompleted, userState)
        End Sub
        
        Private Sub OnGetLotExpensesOperationCompleted(ByVal arg As Object)
            If (Not (Me.GetLotExpensesCompletedEvent) Is Nothing) Then
                Dim invokeArgs As System.Web.Services.Protocols.InvokeCompletedEventArgs = CType(arg,System.Web.Services.Protocols.InvokeCompletedEventArgs)
                RaiseEvent GetLotExpensesCompleted(Me, New GetLotExpensesCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState))
            End If
        End Sub
        
        '''<remarks/>
        Public Shadows Sub CancelAsync(ByVal userState As Object)
            MyBase.CancelAsync(userState)
        End Sub
        
        Private Function IsLocalFileSystemWebService(ByVal url As String) As Boolean
            If ((url Is Nothing)  _
                        OrElse (url Is String.Empty)) Then
                Return false
            End If
            Dim wsUri As System.Uri = New System.Uri(url)
            If ((wsUri.Port >= 1024)  _
                        AndAlso (String.Compare(wsUri.Host, "localHost", System.StringComparison.OrdinalIgnoreCase) = 0)) Then
                Return true
            End If
            Return false
        End Function
    End Class
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.8.9032.0")>  _
    Public Delegate Sub GetSalesCompletedEventHandler(ByVal sender As Object, ByVal e As GetSalesCompletedEventArgs)
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.8.9032.0"),  _
     System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code")>  _
    Partial Public Class GetSalesCompletedEventArgs
        Inherits System.ComponentModel.AsyncCompletedEventArgs
        
        Private results() As Object
        
        Friend Sub New(ByVal results() As Object, ByVal exception As System.Exception, ByVal cancelled As Boolean, ByVal userState As Object)
            MyBase.New(exception, cancelled, userState)
            Me.results = results
        End Sub
        
        '''<remarks/>
        Public ReadOnly Property Result() As System.Data.DataTable
            Get
                Me.RaiseExceptionIfNecessary
                Return CType(Me.results(0),System.Data.DataTable)
            End Get
        End Property
    End Class
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.8.9032.0")>  _
    Public Delegate Sub GetAdjustementsCompletedEventHandler(ByVal sender As Object, ByVal e As GetAdjustementsCompletedEventArgs)
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.8.9032.0"),  _
     System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code")>  _
    Partial Public Class GetAdjustementsCompletedEventArgs
        Inherits System.ComponentModel.AsyncCompletedEventArgs
        
        Private results() As Object
        
        Friend Sub New(ByVal results() As Object, ByVal exception As System.Exception, ByVal cancelled As Boolean, ByVal userState As Object)
            MyBase.New(exception, cancelled, userState)
            Me.results = results
        End Sub
        
        '''<remarks/>
        Public ReadOnly Property Result() As System.Data.DataTable
            Get
                Me.RaiseExceptionIfNecessary
                Return CType(Me.results(0),System.Data.DataTable)
            End Get
        End Property
    End Class
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.8.9032.0")>  _
    Public Delegate Sub GetInventoryActivityCompletedEventHandler(ByVal sender As Object, ByVal e As GetInventoryActivityCompletedEventArgs)
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.8.9032.0"),  _
     System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code")>  _
    Partial Public Class GetInventoryActivityCompletedEventArgs
        Inherits System.ComponentModel.AsyncCompletedEventArgs
        
        Private results() As Object
        
        Friend Sub New(ByVal results() As Object, ByVal exception As System.Exception, ByVal cancelled As Boolean, ByVal userState As Object)
            MyBase.New(exception, cancelled, userState)
            Me.results = results
        End Sub
        
        '''<remarks/>
        Public ReadOnly Property Result() As System.Data.DataTable
            Get
                Me.RaiseExceptionIfNecessary
                Return CType(Me.results(0),System.Data.DataTable)
            End Get
        End Property
    End Class
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.8.9032.0")>  _
    Public Delegate Sub GetTodayInvoicesCompletedEventHandler(ByVal sender As Object, ByVal e As GetTodayInvoicesCompletedEventArgs)
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.8.9032.0"),  _
     System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code")>  _
    Partial Public Class GetTodayInvoicesCompletedEventArgs
        Inherits System.ComponentModel.AsyncCompletedEventArgs
        
        Private results() As Object
        
        Friend Sub New(ByVal results() As Object, ByVal exception As System.Exception, ByVal cancelled As Boolean, ByVal userState As Object)
            MyBase.New(exception, cancelled, userState)
            Me.results = results
        End Sub
        
        '''<remarks/>
        Public ReadOnly Property Result() As System.Data.DataTable
            Get
                Me.RaiseExceptionIfNecessary
                Return CType(Me.results(0),System.Data.DataTable)
            End Get
        End Property
    End Class
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.8.9032.0")>  _
    Public Delegate Sub GetGeneralExpensesCompletedEventHandler(ByVal sender As Object, ByVal e As GetGeneralExpensesCompletedEventArgs)
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.8.9032.0"),  _
     System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code")>  _
    Partial Public Class GetGeneralExpensesCompletedEventArgs
        Inherits System.ComponentModel.AsyncCompletedEventArgs
        
        Private results() As Object
        
        Friend Sub New(ByVal results() As Object, ByVal exception As System.Exception, ByVal cancelled As Boolean, ByVal userState As Object)
            MyBase.New(exception, cancelled, userState)
            Me.results = results
        End Sub
        
        '''<remarks/>
        Public ReadOnly Property Result() As System.Data.DataTable
            Get
                Me.RaiseExceptionIfNecessary
                Return CType(Me.results(0),System.Data.DataTable)
            End Get
        End Property
    End Class
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.8.9032.0")>  _
    Public Delegate Sub GetLotExpensesCompletedEventHandler(ByVal sender As Object, ByVal e As GetLotExpensesCompletedEventArgs)
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.8.9032.0"),  _
     System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code")>  _
    Partial Public Class GetLotExpensesCompletedEventArgs
        Inherits System.ComponentModel.AsyncCompletedEventArgs
        
        Private results() As Object
        
        Friend Sub New(ByVal results() As Object, ByVal exception As System.Exception, ByVal cancelled As Boolean, ByVal userState As Object)
            MyBase.New(exception, cancelled, userState)
            Me.results = results
        End Sub
        
        '''<remarks/>
        Public ReadOnly Property Result() As System.Data.DataTable
            Get
                Me.RaiseExceptionIfNecessary
                Return CType(Me.results(0),System.Data.DataTable)
            End Get
        End Property
    End Class
End Namespace
