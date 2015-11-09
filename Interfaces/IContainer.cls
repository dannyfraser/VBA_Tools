VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Interface for all custom collection classes
'Classes implementing this interface should contain two Private Collections
'One for Items, and another for Keys.
'These should only be accessed using the .Items, .Keys and .Add methods.
'It is also a good idea to add an IUnknown enumeration in the container class so that .Items
'becomes the default member, allowing For Each... loops on the Collection class

'This class has default member attributes set using a text editor.

Private this As ContainerFields
Private Type ContainerFields
    Items As Collection
    Keys As Collection
End Type


Private Sub Class_Initialize()
    Set this.Items = New Collection
    Set this.Keys = New Collection
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
'    Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = Items.[_NewEnum]
End Function


Property Get Items() As Collection
Attribute Items.VB_UserMemId = 0
'    Attribute Items.VB_UserMemId = 0
    Set Items = this.Items
End Property

Property Get Keys() As Collection
    Set Keys = this.Keys
End Property

Sub Add(Item As Variant, Key As Variant)
    If Not Exists(Key) Then
        Items.Add Item, Key
        Keys.Add Key, Key
    End If
End Sub

Sub Remove(Key As Variant)
    If Exists(Key) Then
        Items.Remove Key
        Keys.Remove Key
    End If
End Sub

Function Count() As Long
    Count = Items.Count
End Function

Function Exists(Key As Variant) As Boolean

    On Error GoTo DoesNotExist
    
    Keys(Key) = Keys(Key)
    Exists = True
    
    Exit Function
    
DoesNotExist:
    Exists = False
    
End Function


