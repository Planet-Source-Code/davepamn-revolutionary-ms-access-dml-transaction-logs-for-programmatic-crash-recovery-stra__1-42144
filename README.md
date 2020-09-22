<div align="center">

## Revolutionary\! MS Access DML Transaction logs for programmatic crash recovery strategies


</div>

### Description

MS Access developers will be amazed at this solution. Smart form requires zero programming on your part. What if you could intercept the MS Access sql statement before any database transaction occurred and record this statement in a journal, would this be valuable to you? Of course it would. What if you could replay the journal DML entries in the event a database crash occurred, would this be of value? If you haven’t solved the transaction problem you’ll love this article. Spread the word through the news groups.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[davepamn](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/davepamn.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |ASP \(Active Server Pages\) , VBA MS Access
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/davepamn-revolutionary-ms-access-dml-transaction-logs-for-programmatic-crash-recovery-stra__1-42144/archive/master.zip)





### Source Code

<pre>
Author: David Nishimoto
davepamn@relia.net
<a href="http://www.listensoftware.com">Listen Software Solutions</a>
<a href="http://www.listensoftware.com/hrxp/ds.asp">Design Specification Extreme Pro (Web app using Revolutionary Form</a>
Title: MS Access DML Transaction logs for programmatic crash recovery strategies.
MS Access Revolutionary Forms.
<a href=http://www.listensoftware.com/smartform.zip>Download Sample</a>
Start the movement spread the word
</pre>
<p>
What if you could intercept all MS Access sql statements before any database transaction occurred and record these statements in a journal, would this be valuable to you? Of course it would. What if you could replay the journal DML entries in the event a database crash occurred, would this be of value? Of course, information is money. If you haven’t solved the transaction problem you’ll love this article.
<p>
You may use the <b>Revolutionary Form</b> ideas and code for all non-commerical applications with the stipulation of giving recognition for this technology to Mr. David Nishimoto. Additionally, more robust solutions exist, if you choose to implement <b>Revolutionary Form</b> technology to support a commercial application. David Nishimoto can provide consulting to assist in your implementation. If you choose to use <b>Revolutionary Form</b> technology commercial, Licensing and support are required between your company and David Nishimoto. Enough said, let me tell you about <b>Revolutionary Form</b> technology.
<p>
This article explains how to build a Data Manipulation Language (DML) transaction log in the event the MS Access Database becomes corrupt so data can be recovered. If a disaster occurs the MS Access Schema is restored and all transactions in the DML transaction log are applied. It would be possible with DML transaction logs to create a point in time recovery also speeding up the down time of the crashed database.
<p>
Let me start with the high level solution. First, you need to have a database schema backup. In the event the database becomes non-recoverable then restore the schema, apply the DML transaction log, and bring the database back online.
<p>
Ok, there are some gaps. First MS Access does not have a DML transaction log and second a utility to roll forward the transactions does not exist. So I show you how to build it.
<p>
I decided the solution must be a loosely couple interface existing between what I call a <b>Revolutionary Form</b> and the database. The <b>Revolutionary Form</b> must contain only unbound controls. No bound data except for display or calculation! Use recordset technology to get lookup information from other tables for display fields, but remember DML transaction apply with one table per <b>Revolutionary Form</b>. Data Grid controls are data view only within the <b>Revolutionary Form</b>. No bound data transaction will be recorded in the DML transaction log.
<p>
The <b>Revolutionary Form</b> paves the way for the <b>Revolutionary Form</b> to pass all form information to a data store class. The <b>Revolutionary Form</b> assembles an XML structure from required text boxes and combo controls. Once the XML package has been assembled the <B>Revolutionary Form</b> transmits the package to the data store class for processing. The data store parses the xml package and converts it into an sql statement. Transactions can be either used to manipulate data in the database or extract records package as xml back to the <b>Revolutionary Form</b>. Additionally, the SQL can be stored in a DML transaction ASCII log.
<p>
The Data Field Rules define naming convention for two types of data fields: Required and Primary fields. All other controls on the form are considered either display fields or calculated fields.
<p>
Primary keys represent unique data for a specific table. For example, lbxPNProjectId lbx(1-3) represent the control type (listbox or textbox) other types of control have not been implemented. P (4 character position) - Primary Key used to find the record for update, delete, and select operations. N (5 character position) specifies the data type 1. N-Numeric, D-Date, S-String
<pre>
	<b>Rule 1: A Revolutionary Form may have one and only one data table</b>
	<b>Rule 2: Developers Coding</b>
</pre>
<p>
MS Access developers will be amazed at this solution. Revolutionary form requires zero programming on their part. Play the game by using my field name convention for Revolutionary forms and rules so Revolutionary form can capture dml transactions correctly. Revolutionary forms are for data entry screens only.
Just copy and paste my code for add, update, delete, combo selection, form load, and log preview event (later is recovery). (Highlighted below)
<Pre>
	<b>Sample “Asset” table fields</b>
	Assetid (autonumber)
	Businessunit (number)
	Account (number)
	Serialno (text)
	Invoice (number)
	Purchasedate (date)
	Title (text)
	<b>Rule 2: Developers Coding</b>
	A table can only have a one field primary key and that field must be an autonumber
	<b> Revolutionary Form Fields (name and types></b>
	CboPNAssetid (Primary Key and Numeric)
	TxtRNBusinessUnit (Required Numeric)
	TxtRNAccount 	 (Required Numeric)
	TxtRSSerialno (Required String)
	TxtRNInvoice (Required Numeric)
	TxtRDPurchaseDate (Required Date)
	TxtRSTitle (Required String)
	<b>Revolutionary Form Properties (MS Access Form Properties)</b>
	Single Form
	Allow Edit (Y)
	Allow Deletes (N)
	Allow Additions (N)
	Record Selectors (N)
	Navigation Buttons (N)
	Divided Lines (N)
	<b>XML Assemble display field (Shows the assembled XML package)</b>
 txtFieldList (Hide after testing)
	<b>Text Field for the database table name</b>
 DsTable (set the default value to asset meaning the asset table)
 <b>The DML Transaction Storage location and name</b>
 C:\transactionlog.dat
 <b>Future enhancements</b>
 Support for radio and checkbox data input
</pre>
<p>
The Data Store is a MS Access class which uses the Field Specification to build the sql criteria for one of the follow data manipulation commands: <b>update, delete, and select</b>
<p>
A table contains a grouped list of fields. The Revolutionary form associates one table to a Revolutionary form.
<p>
Required fields are really database fields. If a field is not a required field than it is either a :
	<ul>
	<ol>Display Field
 <ol>A Calculated Field. These fields will be not be sent to the data store class.
 </ul>
<p>
The Revolutionary Form builds an XML package by iterating through all the controls within the Revolutionary form. Each control is examined to determine if it is a text control. Upon qualifying it becomes an attribute name and value pair in the xml structure.
<p>
Upon Revolutionary form request the data store packages an xml package containing one or more data elements representing a data table record. Each xml record is extract by using the Document Object Model (DOM) and mapped to an Revolutionary Form field based on its data name. Field naming convention must follow the data field name rules.
<p>
The data store can perform a number of actions, so of which will return an xml package back to the Revolutionary form.
<pre>
	<b>Add</b> (insert dml into the table)
	<b>Update</b> (update dml for a table using the primary field as criteria)
	<b>Delete</b> (delete dml for a table using the primary field as criteria)
	<b>Select-Single</b> (returns a single xml record using the primary field as criteria)
	<b>Select-Criteria</b> (returns one or more xml record using up to four fields for criteria)
</pre>
The Revolutionary form goal is to increase development speed by creating reusable code. The second goal is to create SQL transactions for each form and store these transactions in a transaction log for recovery.
Data Store Class
<b>The data store class AssembleXML Extract</b>
<pre>
<b>Public Function assembleXML(frmForm, sAction)</b>
 Dim sBuffer
 sBuffer = "<fields><data "
 sBuffer = sBuffer + " process='" + sAction + "' "
 sBuffer = sBuffer + " table='" + frmForm.dsTable.Value + "' "
 For i = 0 To frmForm.Controls.Count - 1
 If frmForm.Controls(i).ControlType = acTextBox _
 Or frmForm.Controls(i).ControlType = acComboBox _
 Or frmForm.Controls(i).ControlType = acListBox Then
 If frmForm.Controls(i).Name <> "process" _
 And frmForm.Controls(i).Name <> "txtFieldList" Then
 If UCase(Mid(frmForm.Controls(i).Name, 4, 1)) = "R" Or UCase(Mid(frmForm.Controls(i).Name, 4, 1)) = "P" Then
  If sBuffer = "" Then
  sBuffer = sBuffer + frmForm.Controls(i).Name + "='" + "" + frmForm.Controls(i).Value + "'"
  Else
  If IsNull(frmForm.Controls(i).Value) Then
  sBuffer = sBuffer + " " + frmForm.Controls(i).Name + "=''"
  Else
  sBuffer = sBuffer + " " + frmForm.Controls(i).Name + "='" + "" + frmForm.Controls(i).Value + "'"
  End If
  End If
 End If
 End If
 Debug.Print sBuffer
 End If
 Next
 sBuffer = sBuffer + " /></fields>"
 assembleXML = sBuffer
<b>end Function</b>
</pre>
<h3>The Revolutionary form calls to the Data Store Class</h3>
<pre>
Option Compare Database
Option Explicit
Private objDataStore As New cDataStore
<b>Private Sub cboPNAssetId_Click()</b>
 Call objDataStore.mapDataStore(Me, "Select-Single")
<b>End Sub</b>
<hr>
<b>Private Sub cmdAdd_Click()</b>
 Dim sql
 Me.txtFieldList.SetFocus
 Me.txtFieldList.Value = objDataStore.assembleXML(Me, "Add")
 sql = objDataStore.ProcessDML(txtFieldList.Text)
 'msgbox sql
<b>End Sub</b>
<hr>
<b>Private Sub cmdDelete_Click()</b>
 Dim sql
 Me.txtFieldList.SetFocus
 Me.txtFieldList.Value = objDataStore.assembleXML(Me, "Delete")
 sql = objDataStore.ProcessDML(txtFieldList.Text)
 Call objDataStore.loadSelection(Me, cboPNAssetId, "assetid", "title", "title", "", "", "", "")
<b>End Sub</b>
<hr>
<b>Private Sub cmdUpdate_Click()</b>
Dim sql
 Me.txtFieldList.SetFocus
 Me.txtFieldList.Value = objDataStore.assembleXML(Me, "Update")
 sql = objDataStore.ProcessDML(txtFieldList.Text)
 'msgbox sql
<b>End Sub</b>
<hr>
<b>Private Sub Form_Load()</b>
 Call objDataStore.loadSelection(Me, cboPNAssetId, "assetid", "title", "title", "", "", "", "")
<b>End Sub</b>
</pre>
<h3>data store class code</h3>
<pre>
Option Compare Database
Option Explicit
Dim sXML
Dim doc
Dim oNode
Dim oAttribute
Dim oData
Dim oFieldNameKeys
Dim i
Dim sql
Dim rs
Dim sProcess
Dim sTableName
Dim oFields
Dim oField
Dim sKeyFieldName
Dim sDisplayFieldName
Dim sCriteria1
Dim sCriteria2
Dim sCriteria3
Dim sCriteria4
Dim sSortFieldName
Dim sUserName
Dim sPassword
<b>Public Function ProcessDML(sXMLPackage)</b>
 Set oData = CreateObject("scripting.dictionary")
 Set doc = CreateObject("Microsoft.XMLDOM")
 doc.async = False
 doc.loadXML sXMLPackage
 Set oNode = doc.selectSingleNode("//data")
 sCriteria1 = ""
 sCriteria2 = ""
 sCriteria3 = ""
 sCriteria4 = ""
 If IsNull(oNode) = False Then
 For Each oAttribute In oNode.Attributes
 If oAttribute.Name = "process" Then
 sProcess = oAttribute.Value
 ElseIf oAttribute.Name = "table" Then
 sTableName = oAttribute.Value
 ElseIf oAttribute.Name = "key_field_name" Then
 sKeyFieldName = oAttribute.Value
 ElseIf oAttribute.Name = "display_field_name" Then
 sDisplayFieldName = oAttribute.Value
 ElseIf oAttribute.Name = "sort_field_name" Then
 sSortFieldName = oAttribute.Value
 ElseIf oAttribute.Name = "criteria1" Then
 sCriteria1 = oAttribute.Value
 ElseIf oAttribute.Name = "criteria2" Then
 sCriteria2 = oAttribute.Value
 ElseIf oAttribute.Name = "criteria3" Then
 sCriteria3 = oAttribute.Value
 ElseIf oAttribute.Name = "criteria4" Then
 sCriteria4 = oAttribute.Value
 ElseIf oAttribute.Name = "username" Then
 sUserName = oAttribute.Value
 ElseIf oAttribute.Name = "password" Then
 sPassword = oAttribute.Value
 Else
 oData.Add oAttribute.Name, oAttribute.Value
 End If
 Next
 oFieldNameKeys = oData.Keys
 End If
 If sProcess = "Add" Then
 sql = BuildInsertSQL()
 CurrentDb.Execute sql
 ProcessDML = sql
 Call RecordDMLTransaction(sql)
 ElseIf sProcess = "Update" Then
 sql = BuildUpdateSQL()
 CurrentDb.Execute sql
 ProcessDML = sql
 Call RecordDMLTransaction(sql)
 ElseIf sProcess = "Delete" Then
 sql = BuildDeleteSQL()
 CurrentDb.Execute sql
 ProcessDML = sql
 Call RecordDMLTransaction(sql)
 ElseIf sProcess = "Security" Then
 sql = "select * from " & sTableName
 sql = sql & " where ucase(username)='" & sUserName & "'"
 sql = sql & " and ucase(password)='" & sPassword & "'"
 'MsgBox sql
 Set rs = CurrentDb.OpenRecordset(sql)
 If Not rs.EOF Then
 'Response.write rs(sKeyFieldName)
 End If
 If Not rs Is Nothing Then
 rs.Close
 End If
 Set rs = Nothing
 ElseIf sProcess = "Select-Single" Then
 sql = BuildSelectSQL()
 sXML = XMLDataStore(sql)
 ProcessDML = sXML
 ElseIf sProcess = "Select-Criteria" Then
 sql = "select " & sKeyFieldName & "," & sDisplayFieldName & " from " & sTableName
 If sCriteria1 <> "" Then
 sql = sql & " where " & sCriteria1
 End If
 If sCriteria2 <> "" Then
 sql = sql & " and " & sCriteria2
 End If
 If sCriteria3 <> "" Then
 sql = sql & " and " & sCriteria3
 End If
 If sCriteria4 <> "" Then
 sql = sql & " and " & sCriteria4
 End If
 sql = sql & " order by " & sSortFieldName
 'Response.Write sql
 sXML = XMLDataStore(sql)
 ProcessDML = sXML
 Else
 MsgBox "Not Found"
 End If
<b>End Function</b>
<hr>
<b>Function XMLDataStore(sql)</b>
 Dim sRetXML
 Dim rs
 Set rs = CurrentDb.OpenRecordset(sql)
 sRetXML = "<fields>"
 Do While Not rs.EOF
 sRetXML = sRetXML + "<data "
 Set oFields = rs.Fields
 For Each oField In oFields
 sRetXML = sRetXML & oField.Name & "='" & oField.Value & "' "
 Next
 rs.MoveNext
 sRetXML = sRetXML + "/>"
 Loop
 sRetXML = sRetXML & "</fields>"
 If Not rs Is Nothing Then
 rs.Close
 End If
 Set rs = Nothing
 XMLDataStore = sRetXML
<b>End Function</b>
<hr>
<b>Function BuildDeleteSQL()</b>
 Dim sql
 Dim sPart1
 Dim sCriteria
 sql = " delete * "
 sCriteria = ""
 For i = 0 To oData.Count - 1
 If Mid(oFieldNameKeys(i), 4, 1) = "P" Then
 If sCriteria = "" Then
 sCriteria = " where " & DBFieldName(oFieldNameKeys(i)) & "="
 Else
 sCriteria = sCriteria & " and " & DBFieldName(oFieldNameKeys(i)) & "="
 End If
 sCriteria = sCriteria & DBFieldValue(oFieldNameKeys(i))
 End If
 Next
 sql = sql & sPart1 & " from [" & sTableName & "] " & sCriteria
 BuildDeleteSQL = sql
<b>End Function</b>
<hr>
<b>Function BuildSelectSQL()</b>
 Dim sql
 Dim sPart1
 Dim sCriteria
 sql = " select "
 sPart1 = ""
 sCriteria = ""
 For i = 0 To oData.Count - 1
 If Mid(oFieldNameKeys(i), 4, 1) = "P" Then
 If sCriteria = "" Then
 sCriteria = " where " & DBFieldName(oFieldNameKeys(i)) & "="
 Else
 sCriteria = sCriteria & " and " & DBFieldName(oFieldNameKeys(i)) & "="
 End If
 sCriteria = sCriteria & DBFieldValue(oFieldNameKeys(i))
 Else
 If Mid(oFieldNameKeys(i), 4, 1) = "R" Then
 If sPart1 = "" Then
  sPart1 = sPart1 & DBFieldName(oFieldNameKeys(i))
 Else
  sPart1 = sPart1 & "," & DBFieldName(oFieldNameKeys(i))
 End If
 End If
 End If
 Next
 sql = sql & sPart1 & " from " & sTableName & " " & sCriteria
 BuildSelectSQL = sql
<b>End Function</b>
<hr>
<b>Function BuildUpdateSQL()</b>
 Dim sql
 Dim sPart1
 Dim sCriteria
 sql = " update " & sTableName & " set "
 sPart1 = ""
 sCriteria = ""
 For i = 0 To oData.Count - 1
 If Mid(oFieldNameKeys(i), 4, 1) = "P" Then
 If sCriteria = "" Then
 sCriteria = " where " & DBFieldName(oFieldNameKeys(i)) & "="
 Else
 sCriteria = sCriteria & " and " & DBFieldName(oFieldNameKeys(i)) & "="
 End If
 sCriteria = sCriteria & DBFieldValue(oFieldNameKeys(i))
 Else
 If Mid(oFieldNameKeys(i), 4, 1) = "R" Then
 If sPart1 = "" Then
  sPart1 = sPart1 & DBFieldName(oFieldNameKeys(i)) & "="
 Else
  sPart1 = sPart1 & "," & DBFieldName(oFieldNameKeys(i)) & "="
 End If
 sPart1 = sPart1 & DBFieldValue(oFieldNameKeys(i))
 End If
 End If
 Next
 sql = sql & sPart1 & sCriteria
 BuildUpdateSQL = sql
<B>End Function</b>
<hr>
<b>Function BuildInsertSQL()</b>
 Dim sPart1
 Dim sPart2
 sPart1 = ""
 sPart2 = ""
 For i = 0 To oData.Count - 1
 'Bypass the primary key fields
 If Mid(oFieldNameKeys(i), 4, 1) <> "P" And Mid(oFieldNameKeys(i), 4, 1) = "R" Then
 If sPart1 = "" Then
 sPart1 = sPart1 & DBFieldName(oFieldNameKeys(i))
 Else
 sPart1 = sPart1 & "," & DBFieldName(oFieldNameKeys(i))
 End If
 If sPart2 = "" Then
 sPart2 = sPart2 & DBFieldValue(oFieldNameKeys(i))
 Else
 sPart2 = sPart2 & "," & DBFieldValue(oFieldNameKeys(i))
 End If
 End If
 Next
 BuildInsertSQL = "insert into " & sTableName & _
 " (" & sPart1 & ")" & _
 " values(" & sPart2 & ")"
<b>End Function</b>
<hr>
<b>Function DBFieldName(sElementName)</b>
 'if mid(sElementName,4,1)="P" then
 DBFieldName = "[" & Right(sElementName, Len(sElementName) - 5) & "]"
 'else
 'DBFieldName=right(sElementName,len(sElementName)-4)
 'end if
<b>End Function</b>
<hr>
<b>Function DBFieldValue(sElementName)</b>
 Dim sValue
 Dim sType
 sValue = oData.Item(sElementName)
 If sValue = "" Then
 DBFieldValue = "Null"
 Exit Function
 End If
 'if mid(sElementName,4,1)="P" then
 sType = Mid(sElementName, 5, 1)
 'else
 'sType=mid(sElementName,4,1)
 'end if
 If sType = "S" Then
 DBFieldValue = "'" & sValue & "'"
 ElseIf sType = "D" Then
 DBFieldValue = "#" & sValue & "#"
 ElseIf sType = "N" Then
 DBFieldValue = sValue
 End If
<b>End Function</b>
<hr>
<b>Public Function assembleXML(frmForm, sAction)</b>
 Dim sBuffer
 sBuffer = "<fields><data "
 sBuffer = sBuffer + " process='" + sAction + "' "
 sBuffer = sBuffer + " table='" + frmForm.dsTable.Value + "' "
 For i = 0 To frmForm.Controls.Count - 1
 If frmForm.Controls(i).ControlType = acTextBox _
 Or frmForm.Controls(i).ControlType = acComboBox _
 Or frmForm.Controls(i).ControlType = acListBox Then
 If frmForm.Controls(i).Name <> "process" _
 And frmForm.Controls(i).Name <> "txtFieldList" Then
 If UCase(Mid(frmForm.Controls(i).Name, 4, 1)) = "R" Or UCase(Mid(frmForm.Controls(i).Name, 4, 1)) = "P" Then
  If sBuffer = "" Then
  sBuffer = sBuffer + frmForm.Controls(i).Name + "='" + "" + frmForm.Controls(i).Value + "'"
  Else
  If IsNull(frmForm.Controls(i).Value) Then
  sBuffer = sBuffer + " " + frmForm.Controls(i).Name + "=''"
  Else
  sBuffer = sBuffer + " " + frmForm.Controls(i).Name + "='" + "" + frmForm.Controls(i).Value + "'"
  End If
  End If
 End If
 End If
 'Debug.Print sBuffer
 End If
 Next
 sBuffer = sBuffer + " /></fields>"
 assembleXML = sBuffer
<b>End Function</b>
<hr>
<b>Public Function loadSelection(frmForm, cboSelection, sKeyFieldName,_
sDisplayFieldName, sSortFieldName, sCriteria1, sCriteria2, _
sCriteria3, sCriteria4)</b>
 Dim sPhrase
 Dim sBuffer
 Dim iIndex
 Dim oNode
 Dim doc
 Dim sXML
 Dim oNodes
 sBuffer = "<fields><data "
 sBuffer = sBuffer + " process='Select-Criteria' "
 sBuffer = sBuffer + " table='" + frmForm.dsTable.Value + "' "
 sBuffer = sBuffer + " key_field_name='" + sKeyFieldName + "' "
 sBuffer = sBuffer + " display_field_name='" + sDisplayFieldName + "' "
 sBuffer = sBuffer + " sort_field_name='" + sSortFieldName + "' "
 If (sCriteria1 <> "") Then
 sBuffer = sBuffer + " criteria1='" + sCriteria1 + "' "
 End If
 If (sCriteria2 <> "") Then
 sBuffer = sBuffer + " criteria2=''+sCriteria2+" ' "
 End If
 If (sCriteria3 <> "") Then
 sBuffer = sBuffer + " criteria3=''+sCriteria3+" ' "
 End If
 If (sCriteria4 <> "") Then
 sBuffer = sBuffer + " criteria4=''+sCriteria4+" ' "
 End If
 sBuffer = sBuffer + " /></fields>"
 'Debug.Print sBuffer
 sXML = ProcessDML(sBuffer)
 Set doc = CreateObject("microsoft.xmldom")
 doc.async = 0
 Call doc.loadXML(sXML)
 Set oNodes = doc.selectNodes("//data")
 Dim iLength
 Dim sKey
 Dim sDisplay
 Dim sFieldName
 Dim j
 Dim sFieldValues
 'Clear the List Box
 cboSelection.Value = ""
 sFieldValues = "assetid; title;"
 cboSelection.RowSourceType = "Value List"
 cboSelection.ColumnCount = 2
 'cboSelection.Clear
 For j = 0 To oNodes.length - 1
 Set oNode = oNodes(j)
 For i = 0 To oNode.Attributes.length - 1
 sFieldName = oNode.Attributes(i).Name
 If UCase(sFieldName) = UCase(sKeyFieldName) Then
  sKey = oNode.Attributes(i).Value
  sFieldValues = sFieldValues & sKey & ";"
 ElseIf UCase(sFieldName) = UCase(sDisplayFieldName) Then
  sDisplay = oNode.Attributes(i).Value
  sFieldValues = sFieldValues & sDisplay & ";"
 End If
 Next
 Next
 cboSelection.RowSource = sFieldValues
 cboSelection.ColumnHeads = True
<b>End Function</b>
<hr>
<b>Public Function mapDataStore(frmForm, sAction)</b>
 Dim sPhrase
 Dim poster
 Dim sXML
 Dim doc
 sXML = assembleXML(frmForm, sAction)
 sXML = ProcessDML(sXML)
 Set doc = CreateObject("microsoft.xmldom")
 doc.async = 0
 doc.loadXML (sXML)
 Dim iLength
 Dim sFieldName
 Dim sFieldValue
 Dim sElementName
 Dim j
 Dim i
 Dim k
 Dim oNodes
 Dim oNode
 Set oNodes = doc.selectNodes("//data")
 For j = 0 To oNodes.length - 1
 Set oNode = oNodes(j)
 For i = 0 To oNode.Attributes.length - 1
 sFieldName = UCase(oNode.Attributes(i).Name)
 sFieldValue = oNode.Attributes(i).Value
 For k = 0 To frmForm.Controls.Count - 1
 If frmForm.Controls(k).ControlType = acTextBox Then
  sElementName = UCase(frmForm.Controls(k).Name)
  If InStr(1, sElementName, sFieldName, vbTextCompare) > 0 Then
  frmForm.Controls(k).Value = sFieldValue
  Exit For
  End If
  Debug.Print sElementName & "," & sFieldName
 End If
 Next
 Next
 Next
<b>End Function</b>
<hr>
<b>Public Sub RecordDMLTransaction(sql)</b>
Dim objFSO
Dim ForReading
Dim ForWriting
Dim ForAppending
Dim sFileName
Dim objCurrent
Dim objStream
 Set objFSO = CreateObject("Scripting.FileSystemObject")
 ForReading = 1
 ForWriting = 2
 ForAppending = 8
 sFileName = "c:\transactionlog.dat"
 If objFSO.FileExists(sFileName) = False Then
 Set objStream = objFSO.CreateTextFile(sFileName, ForWriting)
 Else
 Set objStream = objFSO.OpenTextFile(sFileName, ForAppending)
 End If
 objStream.writeLine (sql)
 objStream.Close
<b>End Sub</b>
<hr>
<b>Public Function RunDMLTransaction()</b>
Dim objFSO
Dim ForReading
Dim ForWriting
Dim ForAppending
Dim sFileName
Dim objCurrent
Dim objStream
Dim sBuffer
Dim sql
 Set objFSO = CreateObject("Scripting.FileSystemObject")
 ForReading = 1
 ForWriting = 2
 ForAppending = 8
 sFileName = "c:\transactionlog.dat"
 If objFSO.FileExists(sFileName) = True Then
 Set objStream = objFSO.OpenTextFile(sFileName, ForReading)
 Do While Not objStream.AtEndOfStream
 sql = objStream.ReadLine
 'currentdb.Execute sql
 sBuffer = sBuffer & sql & Chr(13) & Chr(10) & Chr(13) & Chr(10)
 Loop
 objStream.Close
 End If
 RunDMLTransaction = sBuffer
<b>End Function</b>
</pre>

