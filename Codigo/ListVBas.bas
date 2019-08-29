Attribute VB_Name = "ListVBas"
'Set the varibles used by the Program
Public Db As Database
Public Rs As Recordset
Public Ace As String
'
'
'
'

Public Function LoadLv1()
'Function to Load the Data Base into the ListView
'
      
      Dim colNew As ColumnHeader, NewLine As ListItem
     
     ' Clear the ListView control.
      FrmListV.Lv1.ListItems.Clear
      FrmListV.Lv1.ColumnHeaders.Clear
         
         'Setup The ListView Colomn Headers
         Set colNew = FrmListV.Lv1.ColumnHeaders.Add(, , "Nombre", 1500)
         Set colNew = FrmListV.Lv1.ColumnHeaders.Add(, , "Ciudad")
         Set colNew = FrmListV.Lv1.ColumnHeaders.Add(, , "Estado")
         Set colNew = FrmListV.Lv1.ColumnHeaders.Add(, , "CP")
         Set colNew = FrmListV.Lv1.ColumnHeaders.Add(, , "Tf. casa")
         Set colNew = FrmListV.Lv1.ColumnHeaders.Add(, , "Tf. trabajo")


  'Open the Data Base
  Set Db = OpenDatabase(Path1 & "Data.mdb")
  'Open the Recordset
  Set Rs = Db.OpenRecordset(Format$("Contacts"))

         'Move through the Records
         Rs.MoveFirst
         For i = 1 To Rs.RecordCount
         'put the data into the Item
         Set NewLine = FrmListV.Lv1.ListItems.Add(, , Rs.Fields(0))
         'Put the Data into the SubItems
         NewLine.SubItems(1) = Rs.Fields(1)
         NewLine.SubItems(2) = Rs.Fields(2)
         NewLine.SubItems(3) = Rs.Fields(3)
         NewLine.SubItems(4) = Rs.Fields(4)
         NewLine.SubItems(5) = Rs.Fields(5)
        Rs.MoveNext
        Next i
        'Close the Data Base
        Rs.Close: Set Rs = Nothing: Db.Close: Set Db = Nothing



End Function

Public Function UpdateDb()
'Function to update the Data Base
'

  'Open the Data Base
  Set Db = OpenDatabase(Path1 & "Data.mdb")
  'Open the Recordset
  Set Rs = Db.OpenRecordset("Contacts")

  Rs.MoveFirst
  For i = 1 To Rs.RecordCount
   'If the Record = what we put int the varible Ace Then Update it
  If Ace = Rs.Fields(0) Then
  Rs.Edit
  Rs.Fields(0) = FrmEdit.Text1
  Rs.Fields(1) = FrmEdit.Text2
  Rs.Fields(2) = FrmEdit.Text3
  Rs.Fields(3) = FrmEdit.Text4
  Rs.Fields(4) = FrmEdit.Text5
  Rs.Fields(5) = FrmEdit.Text6
  Rs.Update
  Exit For
  Else
  Rs.MoveNext
  End If
  Next i
  
  'Close the Data Base
  Rs.Close: Set Rs = Nothing: Db.Close: Set Db = Nothing
  
  Unload FrmEdit
  'Reload the Edited Data back into the ListView
  LoadLv1
FrmListV.Visible = True

End Function
