Imports System.IO
Imports System.Threading.Tasks
Imports System.Xml

Module Main
    Private filenamearray As New Hashtable From {{"A", "Audi"}, {"C", "Skoda"}, {"L", "Lnf"}, {"P", "Porsche"}, {"S", "Seat"}, {"V", "Volkwagen"}}
    Private header(1)() As String
    Private input As String
    Private output As String
    Private Const cstUSAGE As String = "Usage: tagestab.exe <input file> <output folder>"
    Private Q As Queue(Of String())

    Sub Main()

#Region "HEADER"

        '        header = {New String() {"Commission number", "Make", "Commission Year", "VIN",
        '"Engine number", "Model year VW", "Model Year", "Model Code",
        '    "Model name", "Model group", "Color code", "Upholstery code",
        '    "Equipment codes", "Status", "GI receipt date", "Schedule Week (week year)",
        '"First planned Week (week year)", "Actual planned Week (week year) ", "Production Week (week year)", "Date of manufacture",
        '"Production number", "Order date", "Scheduling confirmation", "Transfer date",
        '"Confirmation date", "GI-transfer order date", "Kz. Sales / Reserved ", "Client identifier",
        '"Typed (Y / N)", "Invoiced (Y / N)", "Vehicle", "Date of sale",
        '"Delivery date", "Dealer takeover", "Composite main operation", "Trader",
        '"PIA indicator for traders", "Dealer name", "Workshop / Agent", "Workshop / Agent-name",
        '"Intermediaries", "Intermediary designation", "Selling price net dealer", "Selling price net customer",
        '"Seller number", "Customer sales date", "Informal", "Title",
        '"First given name", "Surname", "Country code", "Postal code",
        '"Place", "Authorities order number", "Wholesalers", "TS-valid from",
        '"TS Printing date", "TS number", "Registered Guarantee Date", "Guarantee Date VFW",
        '"Road", "Insurance Association Contribution Date", "VWW-invoice date - min", "-min Factory invoice number",
        '"Invoice date -max", "Invoice date - max", "Margin date", "Invoice price date",
        '"Billing Forced date", "Industry-audience", "Industry-text", "Regular",
        '"Year of birth / date", "Venture", "Brand venture", "Brand-text",
        '"Model venture", "Model Text", "Exchange Kz.", "Year venture",
        '"NV / VFW customer", "Seller Title", "Seller", "Seller surname",
        '"Seller first name", "Seller number Vfw", "Valid from Vfw", "Informal Vfw",
        '"Title Vfw", "First name Vfw", "Surname Vfw", "Country code Vfw",
        '"Postcode Vfw", "Place Vfw", "Road Vfw", "Financial year",
        '"Fiscal month", "Delivery zone work", "Zone of the trader", "Current delivery week",
        '"PS", "KW", "Displacement", "Fuel",
        '"Door count", "Transmission", "Drive", "NOVA",
        '"Drive 1 Consumption city", "Drive 1 country of consumption", "Drive 1 consumption ", "CO2 value for tax calculation",
        '"Particles", "Diselpartikelgrenzwert met (Y / N)", "NOx value", "Varia",
        '"Done message - ZP8", "Financing", "Factory price", "Current checkpoint date",
        '"IFA number", "Private telephone", "Telephone company", "Phone Cell",
        '"Order dealers", "Due date Or payment", "Clearing date", "Internal comment",
        '"Fakturencodes", "GI Speditionskey", "Model Group Name", "Referencing (V = vorl customer / K = Customer / F = Billing / T = type certificate.)",
        '"Down-payment as amount", "Base price for gross payment", "Invoice gross final payment", "CIP model price",
        '"CIP Price Optional extras / color / cushion", "CIP total price", "Dealer comment (FKOMART = 6)", "Difference temp. customer to end customer in days",
        '"Variant", "Destination location", "Stand Day Trader", "Leave-arrived at Hdl",
        '"1. Customers date ", "Date end user", "Customs endorsement:   Customs date", "Status of Charge",
        '"Status of clearance text", "FOB price customs invoice", "Factory Distributor", "File prohibition on disclosure (J = data may Not BE FURTHER / N = DATA MuST BE FURTHER, IN CROSS DEFAULT N)",
        '"E-mail", "Commission no external", "Trim level", "ICO number.",
        '"Purchase price gross GI", "Branch name", "Contact person", "ICO number.",
        '"Accessibility company", "Accessibility Cell", "Accessibility private", "Model grouping category designation",
        '"UID number", "Total billing", "Stand days GI", "Forum",
        '"Billing amount in foreign currency", "Color name long", "TS number / blank TS number", "Version according to factory logic",
        '"GVW", "Maximum load capacity", "Vehicle license", "Final payment status",
        '"Lock number", "Vfw Exit Date", "Currency", "Currency EK",
        '"Shipping data transmitted to the forwarding", "Shipments for vehicle status (01 = arrival / departure 02 =)",
        '"Nummer des Kunden laut SAP", "PBV-Vertragsnummer", "PBV-Laufende Nummer", "Finanzierungsgesellschaft", "Garantiedatum/AaK laut Tagesbericht",
        '"Modelltyp", "PBV-Kz.", "Händlergruppe", "Telefon Fax", "Erreichbarkeit Fax"},
        '                                  New String() {"FZGKOMM", "FZGJAHR", "FABFAB", "FZGFGNR", "FZGMOTN",
        '                                    "FZAMJVW", "FZAMODJ", "MODMODI", "MODMBEZ", "MGTMG",
        '                                    "FARBFARB", "POLPOL", "FZAAUST", "STASTAT", "FZGGIWE",
        '                                    "FZAPLJW", "FZAESGW", "FZAASGW", "FZGPRJW", "FZGPRDA",
        '                                    "FZGPRNR", "FZABEDA", "FZAEIDA", "FZAUEDA", "FZABSDA",
        '                                    "FZGGITA", "FZGVRK", "FZGKKZ", "FZGTS", "FZGFAKT",
        '                                    "FARTART", "FHDVDA", "FHDZUST", "FHDUEDA", "FHDVHBT",
        '                                    "HDLHDL", "PIAKZ", "HDLBEZ", "WSTWST", "WSTBEZ",
        '                                    "FHDVERM", "FHDVEBEZ", "FPRSVKNHDL", "FPRSVKNKND", "FKDVKNR",
        '                                    "FKDGDAV", "ANRANR", "FKDTIT", "FKDVNAM", "FKDNNAM",
        '                                    "PLZLKZ", "PLZPLZ", "FKDORT", "FKDBANR", "FKDGA1",
        '                                    "FTSGDAV", "FTSDDA", "FTSTSNR", "FGADAGADA", "FTSGADAV",
        '                                    "FKDSTR", "FTSVVOTI", "FWRVADA_MIN", "FWRWRNR_MIN", "FAKFDA_MAX",
        '                                    "FAKFNR_MAX", "FZGSPDA", "FZGFPDA", "FHDFAKT", "FSOZBRAN",
        '                                    "SOZBTXT", "FSOZSTAM", "FSOZGEB", "FSOZVORW", "FSOZVMRK",
        '                                    "SOZFTXT", "FSOZVMOD", "SOZMTXT", "FSOZTAU", "FSOZVBJ",
        '                                    "FSOZART", "VKANREDE", "VKTITEL", "VKNACHNAME", "VKVORNAME",
        '                                    "VKDVKNR", "VKDGDAV", "ANRANR_VFW", "VKDTIT_VFW", "VKDVNAM_VFW",
        '                                    "VKDNNAM_VFW", "PLZLKZ_VKD", "PLZPLZ_VKD", "VKDORT", "VKDSTR",
        '                                    "GDATJAHR", "GDATMON", "LIZWZON", "FHDZWLA", "FZGLWON",
        '                                    "TMOTPS", "TMOTKW", "TMOTHUBR", "TREIBANZKZ", "MODTANZ",
        '                                    "GETART", "ANTANTR", "TAUSNOVA", "TAUSVB1", "TAUSVB2",
        '                                    "TAUSVB3", "TAUSCO2", "TAUSPART", "TMOZPGWE", "TAUSNOX",
        '                                    "VARIAS", "FZGZP8", "FKDFIN", "FWRWRPR_SUM", "FZGCPADATE",
        '                                    "FZGIFA", "FKDTNR_PRI", "FKDTNR_FIR", "FKDTNR_HAN", "FZGBHDL", "FAKFDAT_MAX",
        '                                    "FAKAGDA", "KOMM_INTERN", "FAKCODE_ALLE", "FZGGISKEY", "MGTBEZ",
        '                                    "FKDAKZ", "FANZBETR", "FANZBPRS", "FANZRPRS", "CIP_MODEL",
        '                                    "CIP_OPTIONAL", "CIP_TOTAL", "DEALER_COMMENT", "TAGE_ZU_VORLK", "COCVAR",
        '                                    "FLAGORTN", "FZGSTHDL", "FLAGGDAV", "FKDGDAV_MIN", "ENDKUNDE",
        '                                    "FZOLLBDATE", "FZOLLSTAT", "FZOLLSTATTXT", "FZOREFOB", "FHDWHDL",
        '                                    "FKDFREI", "FKDTNR_EMAIL", "FZGKOMMEX", "AUVAUV", "FKDICO",
        '                                    "FPRSEKBGI", "BRABEZ", "KONTAKTPERSON", "VKDICO", "FKDTAKTIV_FIR",
        '                                    "FKDTAKTIV_HAN", "FKDTAKTIV_PRI", "MOGRPBEZ", "FKDUID", "FAKWERT_SUM",
        '                                    "FZGSTGI", "GPTNKB_FORUM", "FAKWERTF_SUM", "MGFABEZF", "BLANKO_TS_NR",
        '                                    "TAUSCVERS", "TAUSHZGW", "TAUSHZNL", "FGADAKFZKZ", "FANZRSTAT",
        '                                    "SPERREN", "FTSVABD", "FPRSVALUTA", "FPRSVALUTAEK", "FVDUEDA",
        '                                    "FVDFSTAT", "FKDNR", "VTRKDNR", "VTRLFNR", "FINGESKZ",
        '                                    "FGADATGBE", "MGTMTYP", "PBVKZ", "BGRPHGRP", "FKDTNR_FAX", "FKDTAKTIV_FAX"}}

#End Region

        If My.Application.CommandLineArgs.Count = 2 Then
            If Not File.Exists(My.Application.CommandLineArgs(0)) Then
                Console.WriteLine("Unable to locate input file.")
                Exit Sub
            End If
            If Directory.Exists(My.Application.CommandLineArgs(1)) Then
                Console.WriteLine("Output folder does not exist. Please create it.")
                Exit Sub
            End If
            If Init() <> 0 Then Exit Sub
            input = My.Application.CommandLineArgs(0)
            output = My.Application.CommandLineArgs(1)
            ProccessFile(input)
        Else
            Console.WriteLine("Ivalid argumets specified.")
            Console.WriteLine()
            Console.WriteLine(cstUSAGE)
        End If



    End Sub

    Private Function Init() As Integer

        Dim status As Integer = 0

        status = LoadHeader()

        Return status

    End Function

    Private Sub ProccessFile(path As String)
        Dim fileHeader As String()
        Dim stpw As New Stopwatch
        Dim line As String()
        Q = New Queue(Of String())

        If Directory.Exists(output) Then

            For Each s As String In Directory.GetFiles(output, "*.csv")

                File.Delete(s)
            Next
        Else
            Directory.CreateDirectory(output)
        End If
        stpw.Restart()

        Using sr As New StreamReader(input)

            'read file header
            fileHeader = sr.ReadLine().Split(",")

            ' if file columns <> header use file header
            If fileHeader.Count <> header(0).Count Then
                Log("File header does not match header file")
                Exit Sub
            End If


            While Not sr.EndOfStream
                line = Split(sr.ReadLine, """,""")

                'For i As Integer = 0 To line.Length - 1
                For i As Integer = 0 To header(1).Count - 1

                    line(i) = line(i).Replace(ControlChars.Quote, "").Replace(";", ".")
                    'FZGPRNR: add 0
                    If i = 20 AndAlso line(i).Length = 6 Then line(i) = $"0{line(i)}"

                    Select Case header(2)(i)
                        Case "date"
                            line(i) = Split(line(i), " ")(0)
                        Case "number"
                            If line(i).Length <> 0 Then
                                If line(i)(0) = "." Then line(i) = $"0{line(i)}"
                                line(i) = Double.Parse(line(i), Globalization.CultureInfo.InvariantCulture)
                            End If
                            line(i) = Replace(line(i), ControlChars.Quote, "")
                        Case Else
                            line(i) = $"""{line(i)}"""
                    End Select

                Next

                'FABFAB = line(2)

                Q.Enqueue(line)

                If Q.Count > 10000 Then
                    Dim t As Task = WriteCSV(Q)

                    'WriteCSV(Q.Where(Function(x As String()) x(2) = "A").ToArray)
                    'WriteCSV(Q.Where(Function(x As String()) x(2) = "V").ToArray)


                    'For Each s As String In filenamearray.Keys
                    '    WriteCSV(Q.Where(Function(x As String()) x(2) = ControlChars.Quote & s & ControlChars.Quote).ToArray)
                    'Next
                    Q.Clear()
                End If



            End While

            'WriteCSV(Q)

        End Using

        stpw.Stop()
        Console.WriteLine($"data proccessed in {stpw.ElapsedMilliseconds} ms")
        Log($"data proccessed in {stpw.ElapsedMilliseconds} ms")

    End Sub

    Private Async Function WriteCSV(q As Queue(Of String())) As Task
        'Dim brand As String '= line(2).Replace(ControlChars.Quote, "")
        Dim sep As String = ";"
        'Dim filename As String '= filenamearray(brand)

        Dim tasks As New List(Of Task)
        Dim destStreams As New List(Of StreamWriter)

        Try
            For Each s As String In filenamearray.Keys
                Dim arr As String()() = q.Where(Function(x As String()) x(2) = ControlChars.Quote & s & ControlChars.Quote).ToArray
                'Dim destStream As New StreamWriter($"{output}\test.csv", True)

                Dim theTask As Task = WriteCSV(arr)
                tasks.Add(theTask)
            Next
            Await Task.WhenAll(tasks)
        Catch ex As Exception

        End Try

    End Function


    Private Async Function WriteCSV(line As String()()) As Task
        If line.Length > 0 Then
            Dim brand As String = line(0)(2).Replace(ControlChars.Quote, "")
            'Console.WriteLine($"brand: {brand}")

            Dim sep As String = ";"
            Dim filename As String = filenamearray(brand)
            'Console.WriteLine($"filename: {filename}")

            Using sw As New StreamWriter($"{output}\{filename}.csv", True)
                For Each l As String() In line
                    Await sw.WriteLineAsync(String.Join(sep, l))
                Next


            End Using
        End If


    End Function


    ''' <summary>
    ''' Log to file 'log.txt'
    ''' </summary>
    ''' <param name="s">String to write to file</param>
    Sub Log(s As String)
        Dim sw As New StreamWriter("log.txt", True)
        sw.WriteLine($"[{Now}] - {s}")
        sw.Close()
    End Sub

    ''' <summary>
    ''' Load header from XML
    ''' </summary>
    ''' <param name="path">Path to file containing header data. Default "header.xml"</param>
    Private Function LoadHeader(Optional path As String = "header.xml") As Integer

        'Check if file exists
        If Not File.Exists(path) Then
            Log($"File not found: {path}")
            'Throw New FileNotFoundException
            Return 1
        End If

        Dim name As New List(Of String)
        Dim altname As New List(Of String)
        Dim type As New List(Of String)

        Dim doc As New XmlDocument
        Dim nl As XmlNodeList

        doc.Load(path)

        nl = doc.SelectNodes("/header/h")

        For i As Integer = 0 To nl.Count - 1
            name.Add(nl(i).Attributes.GetNamedItem("name").Value)
            altname.Add(nl(i).Attributes.GetNamedItem("altname").Value)
            type.Add(nl(i).Attributes.GetNamedItem("type").Value)
        Next

        header = {name.ToArray, altname.ToArray, type.ToArray}

        Return 0
    End Function

End Module
