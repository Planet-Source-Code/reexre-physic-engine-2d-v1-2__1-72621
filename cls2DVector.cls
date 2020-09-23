VERSION 1.0 CLASS
BEGIN
MultiUse = -1 'True
Persistable = 0 'NotPersistable
DataBindingBehavior = 0 'vbNone
DataSourceBehavior  = 0 'vbNone
MTSTransactionMode  = 0 'NotAnMTSObject
END
Attribute VB_Name = "cls2DVector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Author :Roberto Mior
'     reexre@gmail.com
'
'If you use source code or part of it please cite the author
'You can use this code however you like providing the above credits remain intact
'
'
'
'
'--------------------------------------------------------------------------------

'A class to perform vector operations.
'excessive?

Public x As Single
Public y As Single


Public Sub SetV(V As cls2DVector)
    x = V.x
    y = V.y
End Sub

Public Sub AddV(V As cls2DVector)
    x = x + V.x
    y = y + V.y
End Sub
Public Sub SubV(V As cls2DVector)
    x = x - V.x
    y = y - V.y
End Sub
Public Sub MUL(N As Single)
    x = x * N
    y = y * N
End Sub
Public Sub DIV(N As Single)
    x = x / N
    y = y / N
End Sub

Public Function DotV(V As cls2DVector) As Single
    DotV = x * V.x + y * V.y
End Function

'Vector Length (or Magnitude)
Public Function Mag() As Single
    Mag = Sqr(x * x + y * y)
End Function

'Normalize the vector
Public Sub Normalize()
    Dim nMag As Single
    
    nMag = Sqr(x * x + y * y)
    x = x / nMag
    y = y / nMag
End Sub

'Return Normalized Vector
Public Function Normal() As cls2DVector
    Dim nMag As Single
    nMag = Sqr(x * x + y * y)
    Normal.x = x / nMag
    Normal.y = y / nMag
End Function

'Return Perp Vector
Public Function Perpendicular() As cls2DVector
    Set Perpendicular = New cls2DVector
    
    'Perpendicular.x = y
    'Perpendicular.y = -x
    
    Perpendicular.x = -y '
    Perpendicular.y = x '
    
End Function

Public Function Projection(V As cls2DVector) As cls2DVector
    ' def projection(self, vector):
    '        k = (self.dot(vector)) / vector.length()
    '        return k * vector.unit()
    
    Dim K As Single
    
    Set Projection = New cls2DVector
    
    K = (x * V.x + y * V.y) / Sqr(V.x * V.x + V.y * V.y)
    V.Normalize
    V.MUL K
    
    Set Projection = V
End Function

Public Sub Project(V As cls2DVector)
    ' def projection(self, vector):
    '        k = (self.dot(vector)) / vector.length()
    '        return k * vector.unit()
    
    Dim K As Single
    
    K = (x * V.x + y * V.y) / Sqr(V.x * V.x + V.y * V.y)
    V.Normalize
    
    x = V.x * K
    y = V.y * K
    
End Sub


'-------------------------------------------------------------------------------------------------------------

'import Math
'
'def _is_numeric(obj):
'    if isinstance(obj, (int, long, float)):
'        return True
'    Else:
'        return False
'
'Class vector(object):
'
'    def __init__(self, a=0, b=0 ):
'        if _is_numeric(a):
''            #assume two numbers
'            self.x = A
'            self.y = B
'        Else:
'            #assume Vectors/tuples
''            self.x = b[0] - a[0]
'            self.y = b[1] - a[1]
'
'    def __getitem__(self, index):
'        if index == 0:
''            return self.x
'''        elif index == 1:
''            return self.y
'        Else:
'            Raise IndexError
'
'    def __add__(self, other):
'        return Vector(self.x + other.x, self.y + other.y)'
'
'    def __sub__(self, other):
'        return Vector(self.x - other.x, self.y - other.y)'
'
'    def __mul__(self, other):
'Try:
'            other = other - 0
'except:
'            Raise TypeError, "Only scalar multiplication is supported."
'        return Vector( other * self.x, other * self.y )
'
''    def __rmul__(self, other):
'        return self.__mul__(other)
'
'    def __div__(self, other):
'        return Vector( self.x / other, self.y / other )
'
'    def __neg__(self):
'        return Vector(-self.x, -self.y)
'
'    def __abs__(self):
'        return self.length()
'
'    def __repr__(self):
'        Return '(%s, %s)' % (self.x, self.y)
'
'    def __str__(self):
'        Return '(%s, %s)' % (self.x, self.y)
'
'    def dot(self, vector):
'        return self.x * vector.x + self.y * vector.y
'
'    def cross(self, vector):
'        return self.x * vector.y - self.y * vector.x
'
'    def length(self):
'        return math.sqrt( self.dot(self) )
'
'    def perpindicular(self):
'        return Vector(-self.y, self.x)
'
'    def unit(self):
'        return self / self.length()
'
'    def projection(self, vector):
'        k = (self.dot(vector)) / vector.length()
'        return k * vector.unit()
'
'    def angle(self, vector = None):
'        if vector == None:
'            vector = vector(1, 0)
'        return math.acos((self.dot(vector))/(self.length() * vector.length()))
'
'    def angle_in_degrees(self, vector = None):
'        return (self.angle(vector) * 180) /math.pi
'
'
'
