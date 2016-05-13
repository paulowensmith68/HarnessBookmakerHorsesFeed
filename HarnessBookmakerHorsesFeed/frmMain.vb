Public Class frmMain
    Private Sub btnRun_Click(sender As Object, e As EventArgs) Handles btnRun.Click

        Try

            ' ------------------------------
            ' Horse Racing 
            ' ------------------------------
            '
            If My.Settings.ProcessBetfred Then

                gobjEvent.WriteToEventLog("StartProcess:    *----------------------------------------")
                gobjEvent.WriteToEventLog("StartProcess:    *---  Betfred - Updating Horse Racing ---")
                gobjEvent.WriteToEventLog("StartProcess:    *----------------------------------------")
                Dim BookmakerHorsesDbClass As New BookmakerHorsesDbClass()
                BookmakerHorsesDbClass.PollBetfredEvents("WIN", "Betfred", My.Settings.BetfredXMLUrl)
                BookmakerHorsesDbClass = Nothing

            End If

            If My.Settings.ProcessWilliamHill Then

                gobjEvent.WriteToEventLog("StartProcess:    *---------------------------------------------")
                gobjEvent.WriteToEventLog("StartProcess:    *---  William Hill - Updating Horse Racing ---")
                gobjEvent.WriteToEventLog("StartProcess:    *---------------------------------------------")
                Dim BookmakerHorsesDbClass As New BookmakerHorsesDbClass()
                BookmakerHorsesDbClass.PollWilliamHillEvents("WIN", "William Hill", My.Settings.WilliamHillXMLUrl)
                BookmakerHorsesDbClass = Nothing

            End If

            ' Match new odds
            Dim BookmakerHorsesDbMatch1 As New BookmakerHorsesDbClass()
            BookmakerHorsesDbMatch1.MatchHorsesWithBookmakers(7, "WIN")
            BookmakerHorsesDbMatch1 = Nothing

        Catch ex As Exception

            gobjEvent.WriteToEventLog("StartProcess : Process has been killed, general error : " & ex.Message, EventLogEntryType.Error)

        End Try

    End Sub
End Class
