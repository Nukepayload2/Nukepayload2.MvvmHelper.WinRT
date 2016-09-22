Imports System.Reflection

Module SettingViewModelExtension
    Sub LoadProperties(Of T)(Target As T)
        Dim sto = Windows.Storage.ApplicationData.Current.LocalSettings.Values
        For Each prop In GetType(T).GetRuntimeProperties
            If sto.ContainsKey(prop.Name) AndAlso prop.CanWrite Then
                prop.SetValue(Target, sto(prop.Name))
            End If
        Next
    End Sub
    Sub SaveProperties(Of T)(Target As T)
        Dim sto = Windows.Storage.ApplicationData.Current.LocalSettings.Values
        For Each prop In GetType(T).GetRuntimeProperties
            If Not prop.CanRead Then Continue For
            Dim gotv = prop.GetValue(Target)
            If gotv.GetType.ToString.Contains("CoreDispatcher") Then Continue For
            If sto.ContainsKey(prop.Name) Then
                sto(prop.Name) = gotv
            Else
                sto.Add(prop.Name, gotv)
            End If
        Next
    End Sub
    Sub SaveValue(Of T)(Name$, Value As T)
        Dim sto = Windows.Storage.ApplicationData.Current.LocalSettings.Values
        If sto.ContainsKey(Name) Then
            sto(Name) = Value
        Else
            sto.Add(Name, Value)
        End If
    End Sub
    Function LoadValue(Of T)(Name$) As T
        Dim sto = Windows.Storage.ApplicationData.Current.LocalSettings.Values
        If sto.ContainsKey(Name) Then
            Dim v = sto(Name)
            If TypeOf v Is T Then
                Return DirectCast(v, T)
            Else
                Return Nothing
            End If
        End If
        Return Nothing
    End Function
    Sub SaveCredential(ResName$, UserName$, Password$)
        Dim vault = New Windows.Security.Credentials.PasswordVault()
        Try
            Dim cre = vault.FindAllByResource(ResName)
            For Each c In cre
                If c.UserName = UserName Then
                    vault.Remove(c)
                End If
            Next
        Catch ex As Exception
        End Try
        Try
            If Not String.IsNullOrEmpty(Password) Then
                vault.Add(New Windows.Security.Credentials.PasswordCredential(ResName, UserName, Password))
            End If
        Catch ex As Exception
        End Try
    End Sub
    Function LoadCredential$(ResName$, UserName$)
        Dim vault = New Windows.Security.Credentials.PasswordVault()
        Try
            Return vault.Retrieve(ResName, UserName).Password
        Catch ex As Exception
            Return ""
        End Try
    End Function
End Module
