Imports System.Text.RegularExpressions
Imports Newtonsoft.Json

Public Class FrmFlyerDataExtrarctor
    Private Sub BtnExtractData_Click(sender As Object, e As EventArgs) Handles BtnExtractData.Click
        Dim sourceString As String = New System.Net.WebClient().DownloadString("http://flyer.bestbuy.ca/flyers/bestbuy?type=1#!/flyers/bestbuy-weekly?flyer_run_id=130597")
        Dim regex = New Regex("window\[\'flyerData\'\] = (\{.+\});")
        Dim match = regex.Match(sourceString)
        If match.Success Then
            Dim value = match.Groups(1).Value
            Dim flyerData = JsonConvert.DeserializeObject(Of FlyerData)(value)
        End If
    End Sub
End Class
