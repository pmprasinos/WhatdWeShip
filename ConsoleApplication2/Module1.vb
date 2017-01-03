Imports WebfocusDLL

Module Module1

    Sub Main()
        Dim KeepPulling As Boolean = False
        Console.Title = "WhatdWeShip?"
        Dim refe As String = "http://opsfocus01:8080/ibi_apps/Controller?WORP_REQUEST_TYPE=WORP_LAUNCH_CGI&IBIMR_action=MR_RUN_FEX&IBIMR_domain=qavistes/qavistes.htm&IBIMR_folder=qavistes/qavistes.htm%23salesshipmen&IBIMR_fex=pprasino/full_shipreport_qtd.fex&IBIMR_flags=myreport%2CinfoAssist%2Creport%2Croname%3Dqavistes/mrv/shipping_data.fex%2Cshared%2CisFex%3Dtrue%2CrunPowerPoint%3Dtrue&IBIMR_sub_action=MR_MY_REPORT&WORP_MRU=true&&WORP_MPV=ab_gbv&&IBIMR_random=35973&SHIPPED_D="
        Dim OffSetNum As Integer = 14

        Dim WfToday As String = MakeWebfocusDate(Today)
        Dim Tdate As String = WfToday

        Dim wf As New WebfocusDLL.WebfocusModule
        Console.WriteLine("This code was written by MMaeda & PPrasinos feat. GWong & NHansen")
        Console.WriteLine()
        Dim LogInInfo() As String
        Dim writenow As Boolean = False
        If Not wf.IsLoggedIn Then
            Console.Write("Logging into Webfocus...")
            Do Until wf.IsLoggedIn
                LogInInfo = GetUserPasswordandFex()
                wf.LogIn(LogInInfo(0), LogInInfo(1))
            Loop
            Console.WriteLine("DONE")
            Console.WriteLine()
        End If
        refe = Replace(refe, "&IBIMR_sub_action=MR_MY_REPORT", LogInInfo(2))

        If Not wf.IsLoggedIn Then MsgBox("I cant log in to Webfocus" & Chr(13) & "Try again later IDK....")
        Console.WriteLine(Now & "           ")
        WfToday = MakeWebfocusDate(Today)

        Console.Write("  Our shipments today = ")
        Dim j As Object
        j = wf.GetReporth(refe & WfToday & "&l=" & WfToday)
        Try
            Console.WriteLine(OffsetStr("$" & j(1)(0), OffSetNum))
        Catch
            Console.WriteLine(OffsetStr("$0.00", OffSetNum))
        End Try
        Console.Write("Yesterday's shipments = ")
        Dim yest As Date = Now().AddDays(-1)
        Dim WfYesterDay As String = MakeWebfocusDate(Today.AddDays(-1))
        j = wf.GetReporth(refe & WfYesterDay & "&l=" & WfYesterDay)
        Try
            Console.WriteLine(OffsetStr("$" & j(1)(0), OffSetNum))
        Catch
            Console.WriteLine(OffsetStr("$0.00", OffSetNum))
        End Try


        Do While True
            Console.CursorLeft = 0
            If writenow Then
                Console.WriteLine(Now & "           ")

                Console.CursorLeft = 0

                If Not wf.IsLoggedIn Then
                    Console.Write("Logging into Webfocus...")
                    Do Until wf.IsLoggedIn
                        LogInInfo = GetUserPasswordandFex()
                        wf.LogIn(LogInInfo(0), LogInInfo(1))
                    Loop
                    Console.WriteLine("DONE")
                    Console.WriteLine()
                End If

                WfToday = MakeWebfocusDate(Today)

                Console.Write("  Our shipments today = ")
                j = wf.GetReporth(refe & WfToday & "&l=" & WfToday)
                Try
                    Console.WriteLine(OffsetStr("$" & j(1)(0), OffSetNum))
                Catch
                    Console.WriteLine(OffsetStr("$0.00", OffSetNum))
                End Try
            End If

            writenow = True
            Dim dayIndex As Integer = Today.DayOfWeek
            If dayIndex < DayOfWeek.Monday Then
                dayIndex += 7 'Monday is first day of week, no day of week should have a smaller index
            End If
            Dim dayDiff As Integer = dayIndex - DayOfWeek.Monday
            Dim monday As Date = Today.AddDays(-dayDiff)
            Dim WfMonday As String = MakeWebfocusDate(monday)
            'Dim Mday As String =
            Console.Write("            This Week = ")
            j = wf.GetReporth(refe & WfMonday & "&l=" & WfToday)
            Try
                Console.WriteLine(OffsetStr("$" & j(1)(0), OffSetNum))
            Catch
                Console.WriteLine(OffsetStr("$0.00", OffSetNum))
            End Try

            Console.Write("                  QTD = ")
            If Today > #10/2/2016# Then
                j = wf.GetReporth(refe & "10032016" & "&l=" & WfToday)
            Else
                j = wf.GetReporth(refe & "07042016" & "&l=" & WfToday)
            End If


            Try
                Console.WriteLine(OffsetStr("$" & j(1)(0), OffSetNum))
                If Len(j(1)(0)) + 1 >= OffSetNum Then OffSetNum = Len(j(1)(0)) + 1
            Catch
                Console.WriteLine(OffsetStr("$0.00", OffSetNum))
            End Try

            Console.WriteLine()
            For x = 120 To 0 Step -1
                Console.Write("Next report available in " & x & " seconds ")
                Console.CursorLeft = 0
                Threading.Thread.Sleep(1000)
            Next x
            Console.Write("Press any key to run again...                     ")
            Console.CursorLeft = 0
            While Console.KeyAvailable
                Console.ReadKey(True)
            End While
            Dim r As String = ""
            If Not KeepPulling Then
                r = Console.ReadKey(True).Key.ToString
            End If

            If r = "Oem3" Then
                KeepPulling = True
            End If


            If r = "Y" Or r = "y" Then
                Console.Write("Pulling report...                 ")
                Console.CursorLeft = 0
                WfYesterDay = MakeWebfocusDate(Now.Date)

                j = wf.GetReporth(refe & WfYesterDay & "&l=" & WfYesterDay)

                Console.WriteLine("Yesterdays shipments = $")
                Try
                    Console.WriteLine(OffsetStr("$" & j(1)(0), OffSetNum))
                Catch
                    Console.WriteLine(OffsetStr("$0.00", OffSetNum))
                End Try
                Console.WriteLine()
                Console.CursorLeft = 0
                Console.Write("Press any key to run again...")

                While Console.KeyAvailable
                    Console.ReadKey(True)
                End While

                If Not KeepPulling Then
                    Console.ReadKey(True)
                Else
                    Threading.Thread.Sleep(10000)
                End If

            End If
        Loop
    End Sub

    Private Function OffsetStr(Instr As String, Offset As Integer) As String

        Do While Offset > Len(Instr) : Instr = " " & Instr : Loop
        Return Instr
    End Function


    Private Function MakeWebfocusDate(InDate As Date) As String
        Dim sDay As String = InDate.Day
        Dim sMonth As String = InDate.Month
        If Len(sDay) = 1 Then sDay = "0" & sDay
        If Len(sMonth) = 1 Then sMonth = "0" & sMonth
        MakeWebfocusDate = sMonth & sDay & InDate.Year
    End Function


    Private Function GetUserPasswordandFex() As String()
        Dim h As New Random

        Dim Usernames() As String = {"hfaizi", "mreyes", "MALMARAZ", "MARJMAND", "HYANG", "GWONG", "VDELACRUZ", "JTIBAYAN", "JSOLIS", "ASINGH", "GREYES", "JPIMENTEL", "TOSULLIVAN", "MMARTIN", "VLOPEZ", "SLI", "JIMPERIAL", "JHERNANDEZ", "FHARO", "CGOUTAMA", "HGOMEZ", "EGONZALEZ", "CDAROSA"}

        Dim y As Integer = h.Next(0, Usernames.Length)
        Dim ps As String

        Dim FexAdd As String = "&IBIMR_sub_action=MR_MY_REPORT"
        If Usernames(y) <> "pprasinos" Then
            FexAdd = "&IBIMR_sub_action=MR_MY_REPORT&IBIMR_proxy_id=pprasino.htm&"
            ps = ChrW(112) & ChrW(97) & ChrW(115) & ChrW(115) & ChrW(50) & ChrW(48) & ChrW(49) & ChrW(53)
        Else
            ps = ChrW(87) & ChrW(121) & ChrW(109) & ChrW(97) & ChrW(110) & ChrW(49) & ChrW(50) & ChrW(51) & ChrW(45)
        End If
        Debug.Print(Usernames(y))
        Return {Usernames(y), ps, FexAdd}

    End Function


End Module
