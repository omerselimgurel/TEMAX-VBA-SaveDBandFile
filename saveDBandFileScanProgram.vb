Private WithEvents Items As Outlook.Items
Private Sub Application_Startup()
  Dim olApp As Outlook.Application
  Dim objNS As Outlook.NameSpace
  Set olApp = Outlook.Application
  Set objNS = olApp.GetNamespace("MAPI")
  ' default local Inbox
  Set Items = objNS.GetDefaultFolder(olFolderInbox).Items
End Sub
Private Sub Items_ItemAdd(ByVal item As Object)
Dim objOL As Outlook.Application
Dim objMsg As Outlook.MailItem 'Object
Dim objAttachments As Outlook.Attachments
Dim objSelection As Outlook.Selection
Dim i As Long
Dim lngCount As Long
Dim strFile As String
Dim strFolderpath As String
Dim qryScan As String
qryScan = "QA SCANS FROM"
StrCustomer = "_Customer_"


  On Error GoTo ErrorHandler
  Dim Msg As Outlook.MailItem
  If TypeName(item) = "MailItem" Then
    Set Msg = item
    ' ******************
    ' do something here
    ' ******************
    
    
    isqaScan = isContain(Msg.Subject, qryScan)
    
    If isqaScan = True Then
    'MsgBox ("Doğru İçeriyor")
    
    ' Get the path to your My Documents folder
    strFolderpath = CreateObject("WScript.Shell").SpecialFolders(16)
    
    ' Instantiate an Outlook Application object.
    'Set objOL = CreateObject("Outlook.Application")
    
    ' Get the collection of selected objects.
    'Set objSelection = objOL.ActiveExplorer.Selection
    
    ' Set the Attachment folder.
    strFolderpath = strFolderpath & "\Attachments\"
    
    

    ' This code only strips attachments from mail items.
    ' If objMsg.class=olMail Then
    ' Get the Attachments collection of the item.
    Set objAttachments = Msg.Attachments
    lngCount = objAttachments.Count
    strDeletedFiles = ""

    If lngCount > 0 Then

        ' We need to use a count down loop for removing items
        ' from a collection. Otherwise, the loop counter gets
        ' confused and only every other item is removed.

        For i = lngCount To 1 Step -1

            ' Save attachment before deleting from item.
            ' Get the file name.
            strFile = objAttachments.item(i).FileName
            
            If isContain(strFile, StrCustomer) <> True Then
                ' Combine with the path to the Temp folder.
                strFile = strFolderpath & strFile

                ' Save the attachment as a file.
                objAttachments.item(i).SaveAsFile strFile
                
                readXls (strFile)

                ' Delete the attachment.
                'objAttachments.item(i).Delete
            End If

        Next i

 
    End If

    End If
         
  End If
ProgramExit:
  Exit Sub
ErrorHandler:
  MsgBox Err.Number & " - " & Err.Description
  Resume ProgramExit
End Sub
'Read Excel File
Function readXls(ByVal targetFile As String)

Dim objExcel As New Excel.Application
Dim exWb As Excel.Workbook
'targetFile is excel filepath

Set exWb = objExcel.Workbooks.Open(targetFile)

Dim stopCount As Boolean
Dim rowCount As Integer
Dim cell As String
stopCount = False
rowCount = 0
Dim i As Integer
i = 1
'We just Define null not neccessary
cell = "null"

'cell = exWb.Sheets("qryAllScans").Cells(1, 0)

'We Check Check cell row is empty or not and Count Row Count
While stopCount <> True
    If cell = "" Then
        stopCount = True
    Else
        cell = exWb.Sheets("qryAllScans").Cells(i, 1)
        Inc rowCount
        Inc i
        'MsgBox (cell)
    End If
Wend

'MsgBox (rowCount - 2)

Dim real_row_count As Integer
real_row_count = rowCount - 2
Dim k As Integer
Dim excelArr(0 To 25) As Variant



For i = 1 To real_row_count
    For k = 1 To 26
        cell = exWb.Sheets("qryAllScans").Cells(i + 1, k)
        excelArr(k - 1) = cell
    Next k
    writeDatabase (excelArr)
Next i



End Function
'Write Database String Values
Function writeDatabase(ByVal dataArray As Variant)
'MsgBox (dataArray)
Dim StringUrl As String
Dim hReq As Object, JSON As Dictionary

    'MsgBox (dataArray(2))
    'http://localhost/Matematigo%20-%20Logistika/insertQryAllScans.php?action=insertQryScan&PORefID=testporefid&ProductID=testproductid&BarcodeID=testbarcodeid&ScannerID=testscannnerid&Company_Name=testcompanyname&Name_Scanner=testnamescanner&OrderID=testorderid&MachineID=testmachineid&ProductTypeID=testproducttypeid&ActivityID=testactivityid&TblOrder_DateStamp=testdatestamp&UniqueID=testuniqueid&ComputerID=testcomputerid&OrderDetailID=testorderdetail&TblOrderDetail_DateStamp=testtabledetail&TimeStamp=testtimestamp&Remark=testremark&PersonalExitID=testpersonalexit&DateExitStamp=testdateexit&TimeExitStamp=testtimeexit&Expr1=testexp1&SealNr=testsealnr&Comment=testcomment&QACountry=testqacounty&RemarkID=testremark&QAEmail=testqamail
    StringUrl = "http://localhost/Matematigo%20-%20Logistika/insertQryAllScans.php?action=insertQryScan&PORefID=" & dataArray(0) & "&ProductID=" & dataArray(1) & "&BarcodeID=" & dataArray(2) & "&ScannerID=" & dataArray(3) & "&Company_Name=" & dataArray(4) & "&Name_Scanner=" & dataArray(5) & "&OrderID=" & dataArray(6) & "&MachineID=" & dataArray(7) & "&ProductTypeID=" & dataArray(8) & "&ActivityID=" & dataArray(9) & "&TblOrder_DateStamp=" & dataArray(10) & "&UniqueID=" & dataArray(11) & "&ComputerID=" & dataArray(12) & "&OrderDetailID=" & dataArray(13) & "&TblOrderDetail_DateStamp=" & dataArray(14) & "&TimeStamp=" & dataArray(15) & "&Remark=" & dataArray(16) & "&PersonalExitID=" & dataArray(17) & "&DateExitStamp=" & dataArray(18) & "&TimeExitStamp=" & dataArray(19) & "&Expr1=" & dataArray(20) & "&SealNr=" & dataArray(21) & "&Comment=" & dataArray(22) & "&QACountry=" & dataArray(23) & "&RemarkID=" & dataArray(24) & "&QAEmail=" & dataArray(25)
    
    Set hReq = CreateObject("MSXML2.XMLHTTP")
    With hReq
        .Open "POST", StringUrl, False
        .Send
    End With
    
Set hReq = Nothing
    
End Function

Function Inc(ByRef i As Integer)
   i = i + 1
End Function

Function isContain(ByVal i As String, ByVal k As String) As Boolean
    'i is File name string, k is Customer String'
Dim tempNumber As Integer

    For tempNumber = 1 To (Len(i) - Len(k))
    
        If Mid(i, tempNumber, Len(k)) = k Then
            isContain = True
        End If
    
    Next tempNumber
    
End Function
