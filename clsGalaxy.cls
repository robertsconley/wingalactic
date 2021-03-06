VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGalaxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarDirectory As String 'local copy
Private mvarGalaxyName As String 'local copy
Public Property Let GalaxyName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GalaxyName = 5
    mvarGalaxyName = vData
End Property


Public Property Get GalaxyName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GalaxyName
    GalaxyName = mvarGalaxyName
End Property



Public Property Let Directory(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Directory = 5
    mvarDirectory = vData
End Property


Public Property Get Directory() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Directory
    Directory = mvarDirectory
End Property



