Attribute VB_Name = "modVector"
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

Public Function CIRCLE_LINE(P1 As tPoint, P2 As tPoint, C As tPoint, R As Single) As tPoint
    'perform Circle - Line Collision and return Contact Point
    'Get and translated to vb6 from some site... Really don't know how it works.
    'Intersection between Circle and Line (segment) generally should return 2 Points.
    'I suppose this return the middle point between them.
    
    Dim A As New cls2DVector
    Dim B As New cls2DVector
    Dim P As New cls2DVector
    Dim AC As New cls2DVector
    Dim AB As New cls2DVector
    Dim H As New cls2DVector
    Dim CC As New cls2DVector
    
    Dim H2 As Single
    Dim R2 As Single
    
    
    Dim AB2 As Single
    Dim Acab As Single
    Dim T As Single
    
    A.x = P1.x
    A.y = P1.y
    B.x = P2.x
    B.y = P2.y
    
    AC.x = C.x
    AC.y = C.y
    
    CC.x = C.x
    CC.y = C.y
    
    
    AC.SubV A
    
    AB.x = B.x
    AB.y = B.y
    
    AB.SubV A
    
    AB2 = AB.DotV(AB)
    Acab = AC.DotV(AB)
    
    T = Acab / AB2
    
    If (T < 0) Then
        T = 0
    ElseIf (T > 1) Then
        T = 1
    End If
    
    'P = A + t * AB
    P.x = AB.x
    P.y = AB.y
    P.MUL T
    P.AddV A
    
    
    H.x = P.x
    H.y = P.y
    
    H.SubV CC
    
    
    H2 = H.DotV(H)
    R2 = R * R
    
    If (H2 > R2) Then
        CIRCLE_LINE.x = -99
        CIRCLE_LINE.y = -99
    Else
        CIRCLE_LINE.x = P.x
        CIRCLE_LINE.y = P.y
    End If
    
    
    '
End Function




'--------------------------------------------------------------------
'
'
Public Function getREFLECT(Vv As tPoint, Wall As tPoint) As tPoint
    'Function returning the reflection of one vector around another.
    'Get and translated to vb6 from some site... Really don't know how it works.
    '
    'it's used to calculate the rebound of a Vector on another Vector
    'Vector "VV" represents current velocity of a point.
    'Vector "Wall" represent the angle of a wall where the point Bounces.
    '
    'Returns the vector velocity that the point takes after the rebound
    '(Returned values .x, and .y, are reversed)
    
    Dim C1 As Single
    Dim r1 As New cls2DVector
    Dim V As New cls2DVector
    Dim N As New cls2DVector
    Dim Nd As New cls2DVector
    
    V.x = Vv.x
    V.y = Vv.y
    N.x = Wall.x
    N.y = Wall.y
    Nd.x = Wall.x
    Nd.y = Wall.y
    
    N.Normalize
    
    'Vect2 = Vect1 - 2 * WallN * (WallN DOT Vect1)
    C1 = N.DotV(V)
    N.MUL C1
    N.MUL 2
    
    V.SubV N
    
    getREFLECT.x = V.x
    getREFLECT.y = V.y
    
End Function


Public Function MakePoint(x, y) As tPoint
    MakePoint.x = x
    MakePoint.y = y
End Function

Public Function MakeVector(x, y) As cls2DVector
    Set MakeVector = New cls2DVector
    
    MakeVector.x = x
    MakeVector.y = y
End Function

Public Function Normalize(P As tPoint) As tPoint
    Dim Le As Single
    Le = Sqr(P.x * P.x + P.y * P.y)
    Normalize.x = P.x / Le
    Normalize.y = P.y / Le
End Function





'public function LineToLineIntersection(x1_:Number, y1_:Number, x2_:Number, y2_:Number, x3_:Number, y3_:Number, x4_:Number, y4_:Number):Object{
'            var result:Object = {b:false, x:0, y:0};
'            var r:Number, s:Number, d:Number;
'            d = (((x2_-x1_)*(y4_-y3_))-(y2_-y1_)*(x4_-x3_));
'            if(d != 0){
'                r = (((y1_-y3_)*(x4_-x3_))-(x1_-x3_)*(y4_-y3_))/d;
'                s = (((y1_-y3_)*(x2_-x1_))-(x1_-x3_)*(y2_-y1_))/d;
'                if (r >=0 && r <= 1){
'                    if (s >= 0 && s <= 1){
'                        result.b = true;
'                        result.x = x1_ + r * (x2_ - x1_);
'                        result.y = y1_ + r * (y2_ - y1_);
'                    }
'                }
'            }
'            return result;
'        }
Public Function LINE_LINE(L1P1 As tPoint, L1P2 As tPoint, L2P1 As tPoint, L2P2 As tPoint) As tPoint
    Dim R As Double
    Dim S As Double
    Dim D As Double
    
    LINE_LINE.x = 0
    LINE_LINE.y = 0
    
    D = (L1P2.x - L1P1.x) * (L2P2.y - L2P1.y) - (L1P2.y - L1P1.y) * (L2P2.x - L2P1.x)
    
    If D <> 0 Then
        
        R = ((L1P1.y - L2P1.y) * (L2P2.x - L2P1.x) - (L1P1.x - L2P1.x) * (L2P2.y - L2P1.y)) / D
        
        If R >= -0 And R <= 1 Then
            
            S = ((L1P1.y - L2P1.y) * (L1P2.x - L1P1.x) - (L1P1.x - L2P1.x) * (L1P2.y - L1P1.y)) / D
            
            If S >= -0 And S <= 1 Then
                LINE_LINE.x = L1P1.x + R * (L1P2.x - L1P1.x)
                LINE_LINE.y = L1P1.y + R * (L1P2.y - L1P1.y)
                
            End If
            
        End If
        
    End If
    
End Function




'***********************************************************
'***********************************************************
'public static DVec2 CIRCLE_LINE(DVec2 C, double r, Line line)
'{
'    DVec2 A = line.p1;
'    DVec2 B = line.p2;
'    DVec2 P;
'    DVec2 AC = new DVec2( C );
'    AC.sub(A);
'    DVec2 AB = new DVec2( B );
'    AB.sub(A);
'    double ab2 = AB.dot(AB);
'    double acab = AC.dot(AB);
'    double t = acab / ab2;
'
'    if (t < 0.0)
'        t = 0.0;
'    else if (t > 1.0)
'        t = 1.0;
'
'    //P = A + t * AB;
'    P = new DVec2( AB );
'    P.mul( t );
'    P.add( A );
'
'    DVec2 H = new DVec2( P );
'    H.sub( C );
'    double h2 = H.dot(H);
'    double r2 = r * r;
'
'    if(h2 > r2)
'        return null;
'    Else
'        return P;
'}
'
''
'----------------------------------------------------------------------------
'public static DVec2 lineIntersectsLine( DVec2 v1, DVec2 v2, DVec2 v3, DVec2 v4 )
'{
'    double denom = ((v4.y - v3.y) * (v2.x - v1.x)) - ((v4.x - v3.x) * (v2.y - v1.y));
'    double numerator = ((v4.x - v3.x) * (v1.y - v3.y)) - ((v4.y - v3.y) * (v1.x - v3.x));
'
'    double numerator2 = ((v2.x - v1.x) * (v1.y - v3.y)) - ((v2.y - v1.y) * (v1.x - v3.x));
'
'    if ( denom == 0.0f )
'    {
'        if ( numerator == 0.0f && numerator2 == 0.0f )
'        {
'            return null; //COINCIDENT;
'        }
'        return null; //PARALLEL;
'    }
'    double ua = numerator / denom;
'    double ub = numerator2/ denom;
'
'    if(ua >= 0.0f && ua <= 1.0f && ub >= 0.0f && ub <= 1.0f)
'    {
'        return add( v1, mul( sub(v2, v1), ua ) );
'    }
'    Else
'        return null;
'}


'----------------------------------------------------------------------------
'public class DVec2
'{
'    public double x;
'    public double y;
'
'    public DVec2(double x, double y)
'    {
'        this.x = x;
'        this.y = y;
'    }
'
'    Public DVec2()
'    {
'        x = 0.0;
'        y = 0.0;
'    }
'
'    public void set(DVec2 p)
'    {
'        x = p.x;
'        y = p.y;
'    }
'
'    public void add(DVec2 p)
'    {
'        x += p.x;
'        y += p.y;
'    }
'
'    public void sub(DVec2 p)
'    {
'        x -= p.x;
'        y -= p.y;
'    }
'
'    public DVec2 mul(double v)
'    {
'        x *= v;
'        y *= v;
'        return this;
'    }
'
'    public void div(double v)
'    {
'        x /= v;
'        y /= v;
'    }
'
'    public void normalize()
'    {
'        div(mag());
'    }
'
'    public double dot( DVec2 v )
'    {
'        return x*v.x + y*v.y;
'    }
'
'    public double mag()
'    {
'        return Math.sqrt(Math.pow(x,2) + Math.pow(y,2));
'    }
'
'    public String toString()
'    {
'        return "x: " + x + ", y: " + y;
'    }
'
'}
'
'
'



'--------------------------------------------------------------------
'Function returning the reflection of one vector around another.
'
'public static Vector3d getREFLECT(Vector3d v, Vector3d n)
'{
'    double c1 = -n.dot( v );
'    Vector3d r1 = new Vector3d();
'    r1.set(n);
'    r1.scale(c1);
'    r1.scale(2.0);
'    r1.add(v);
'    return r1;
'}





'/* Taken from Robert Sedgewick, Algorithms in C++ */
'
'/*  returns whether, in traveling from the first to the second
'    to the third point, we turn counterclockwise (+1) or not (-1) */
'
'int ccw( Point p0, Point p1, Point p2 )
'{
'    int dx1, dx2, dy1, dy2;
'
'    dx1 = p1.x - p0.x; dy1 = p1.y - p0.y;
'    dx2 = p2.x - p0.x; dy2 = p2.y - p0.y;
'
'    if (dx1*dy2 > dy1*dx2)
'        return +1;
'    if (dx1*dy2 < dy1*dx2)
'       return -1;
'   if ((dx1*dx2 < 0) || (dy1*dy2 < 0))
'        return -1;
'    if ((dx1*dx1 + dy1*dy1) < (dx2*dx2 + dy2*dy2))
'        return +1;
'    return 0;
'}



'//line segment goes from (x1,y1) to (x2,y2)
'2   //the test point is at (x,y)
'3   float A = x - x1;//vector from one end point to the test point
'4   float B = y - y1;
'5   float C = x2 - x1;//vector from one end point to the other end point
'6   float D = y2 - y1;
'7
'8   float dot = A * C + B * D;//some interesting math coming from the geometry of the algorithm
'9   float len_sq = C * C + D * D;
'10  float param = dot / len_sq;
'11
'12  float xx,yy;//the coordinates of the point on the line segment closest to the test point
'13
'14  //the parameter tells us which point to pick
'15  //if it is outside 0 to 1 range, we pick one of the endpoints
'16  //otherwise we pick a point inside the line segment
'17  if(param < 0)
'18  {
'19  xx = x1;
'20  yy = y1;
'21  }
'22  else if(param > 1)
'23  {
'24  xx = x2;
'25  yy = y2;
'26  }
'27  Else
'28  {
'29  xx = x1 + param * C;
'30  yy = y1 + param * D;
'31  }
'32
'33  float dist = dist(x,y,xx,yy);//distance from the point to the segment''
'
''





'Get the normal vector of that line plus any point in that line.'
'
'S = <point1 - point_in_line,normal>*<point2 - point_in_line,normal>
'
'Being <a,b> the dot product.
'
'S < 0 => both are on the same side of the line (angle between them is in the interval [0,PI/2)).
'S = 0 => at least one is on the line.
'S < 0 => both are on different sides of the line (angle between them is in the interval (PI/2,PI]).
'
'
'bool points_on_same_side_of_line(const Vector2d &p1, const Vector2d &p2,
'                                 const Vector2d &p_line, const Vector2d &normal)
'{
'    return normal.dot(p1 - p_line)*normal.dot(p2 - p_line) > 0.0f;
'}
'
'



'function PointOnLine2D(x1,y1, x2,y2, x3,y3:double):boolean;
'begin
'result:=((x2 - x1) * (y3 - y1)) - ((x3 - x1) * (y2 - y1)) < 0.0000001;
'end;
