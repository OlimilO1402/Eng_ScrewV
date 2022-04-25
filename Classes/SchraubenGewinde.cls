VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Schraubengewinde"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'should this really be a class?
'or is it better just one Select Case inside each of the 2 Properties of Schraube?
'ist diese Entscheidung irgendwie beeinflusst? Möglicherweise von der verwendeten Programmiersprache?
'
'
'Alle Angaben in [mm]

'Regelgewinde
'Gewinde | Steigung | Durchmesser Kernloch/Seele
'M 6     |   1      |     5
'M 8     |   1,25   |     6,8
'M10     |   1,5    |     8,5
'M12     |   1,75   |    10,2
'M16     |   2      |    14
'M20     |   2,5    |    17,5
'M24     |   3      |    21
'M30     |   3,5    |    26,5


'Feingewinde
'Gewinde | Steigung | Durchmesser Kernloch/Seele

' M6     |   0,75   |     5,2

' M8     |   0,75   |     7,2
' M8     |   1      |     7

'M10     |   0,75   |     9,2
'M10     |   1      |     9
'M10     |   1,5    |     8,8

'M12     |   1      |    11
'M12     |   1,25   |    10,8
'M12     |   1,5    |    10,5

'M16     |   1      |    15
'M16     |   1,5    |    14,5

'M20     |   1      |    19
'M20     |   1,5    |    18,5
'M20     |   2      |    18

'M24     |   1      |    23
'M24     |   1,5    |    22,5
'M24     |   2      |    22

'M30     |   1      |    29
'M30     |   1,5    |    28,5
'M30     |   2      |    28

