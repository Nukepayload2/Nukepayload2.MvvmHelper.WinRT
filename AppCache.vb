Imports System.Text
Imports Newtonsoft.Json

Public Class AppCache
    Public Shared Async Function LoadAsync(Data As Object, <CallerFilePath> Optional CallerFilePath$ = "") As Task
        Dim localUserData = Await OpenDataFileAsync(CallerFilePath)
        Using strm = Await localUserData.OpenStreamForReadAsync, dr = New StreamReader(strm, Encoding.Unicode)
            Dim txt = Await dr.ReadToEndAsync
            JsonConvert.PopulateObject(txt, Data)
        End Using
    End Function
    Public Shared Async Function SaveAsync(Data As Object, <CallerFilePath> Optional CallerFilePath$ = "") As Task
        Dim localUserData = Await OpenDataFileAsync(CallerFilePath)
        Using strm = Await localUserData.OpenStreamForWriteAsync, dw = New StreamWriter(strm, Text.Encoding.Unicode)
            strm.SetLength(0)
            Await dw.WriteAsync(JsonConvert.SerializeObject(Data))
        End Using
    End Function

    Private Shared Async Function OpenDataFileAsync(CallerFilePath As String) As Task(Of Windows.Storage.StorageFile)
        Dim CallerFileName = Path.GetFileName(CallerFilePath)
        Dim localUserData = Await Windows.Storage.ApplicationData.Current.LocalFolder.CreateFileAsync(CallerFileName + ".json", Windows.Storage.CreationCollisionOption.OpenIfExists)
        Return localUserData
    End Function
End Class