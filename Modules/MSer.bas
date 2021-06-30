Attribute VB_Name = "MSer"
Option Explicit
'Modul zum Serialisieren
Private m_FSO As cFSO

Private m_IDs As Collection

'Public Sub JSONSerialize(FSO As cFSO, aPFN As String)
'    Dim myArchive As cCollection: Set myArchive = New_c.JSONObject
'    'myArchive.prop("Circle.Radius") = CStr(2.5)
'    'myArchive.prop("Circle.CenterX") = CStr(15)
'    'myArchive.prop("Circle.CenterY") = CStr(25)
'
'    Dim s As String
'    s = myArchive.SerializeToJSONString
'    'Set m_FSO = New_c.FSO
'    FSO.WriteTextContent aPFN, s
'
'End Sub

Public Function JSONSerializeVBObject(Obj As Object) As String
Dim Item As Object, JSONObj As cCollection, Props As cProperties, Prop As cProperty

  Set JSONObj = New_c.JSONObject 'create a new target-container (a JSON-capable cCollection)
  
  Set Props = New_c.Properties 'create an instance of the RC5-Properties-Enumerator ...
      Props.BindTo Obj, True   '... and bind it to the Obj-Argument we got handed
  
  For Each Prop In Props 'now enumerate the Properties, to transfer them into the JSONObj
    If Not IsObject(Prop.Value) Then 'it's a normal Value we can just assign to the JSONObj
      JSONObj.Prop(Prop.Name) = Prop.Value
    ElseIf TypeOf Prop.Value Is Collection Then 'it's a VBA.Collection, and so...
      JSONObj.Prop(Prop.Name) = New_c.JSONArray '...we'll treat its content like a JSON-Array
      For Each Item In Prop.Value 'enumerate the Collection-Items (which need to be Objects)
        JSONObj.Prop(Prop.Name).Add New_c.JSONDecodeToCollection(JSONSerializeVBObject(Item))
      Next
    Else 'it's a normal Object, which is not "a List-Object" (as e.g. a VBA.Collection)
      JSONObj.Prop(Prop.Name) = New_c.JSONDecodeToCollection(JSONSerializeVBObject(Prop.Value))
    End If
  Next
  
  JSONSerializeVBObject = JSONObj.SerializeToJSONString
End Function

Public Function JSONDeSerializeVBObject(Obj As Object, JSONObj As cCollection) As Object
Dim Item As Object, Props As cProperties, Prop As cProperty

  Set Props = New_c.Properties 'create an instance of the RC5-Properties-Enumerator ...
      Props.BindTo Obj, True   '... and bind it to the Obj-Argument we got handed
  
  For Each Prop In Props 'now enumerate the Properties, to be able to copy them from JSONObj
    If Not IsObject(Prop.Value) Then 'it's a normal Value we can just assign from the JSONObj
      Prop.Value = JSONObj.Prop(Prop.Name)
      
    ElseIf TypeOf Prop.Value Is Collection Then 'it's a VBA.Collection, and so...
      For Each Item In JSONObj.Prop(Prop.Name) 'enumerate the JSON-Items of the JSON-Array
        Prop.Value.Add JSONDeSerializeVBObject(CreateInstanceByPropName(Prop.Name), Item)
      Next
    Else 'it's a normal Object, which is not "a List-Object" (as e.g. a VBA.Collection)
      Set Prop.Value = JSONDeSerializeVBObject(CreateInstanceByPropName(Prop.Name), JSONObj.Prop(Prop.Name))
    End If
  Next

  Set JSONDeSerializeVBObject = Obj
End Function

Private Function CreateInstanceByPropName(PropName As String) As Object
    Select Case LCase$(PropName)
    Case "Schraube":        Set CreateInstanceByPropName = New Schraube
    Case "Schraubengruppe": Set CreateInstanceByPropName = New Schraubengruppe
    Case "Schraubenloch":   Set CreateInstanceByPropName = New Schraubenloch
    
    'Case "phonenumbers": Set CreateInstanceByPropName = New cPhone
    'Case "children":     Set CreateInstanceByPropName = New cPerson
    End Select
End Function
