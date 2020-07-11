Option Explicit On
Option Strict On

Public NotInheritable Class DBConnString

    'VPN  7.191.194.14    
    'Public Shared strConn2 As String = "Data Source=7.49.100.224\SQLEXPRESS;Initial Catalog=DB2012;User ID=sa;Password=$y$05000"
    '  DataBaseTest 

    'Public Shared strConn2 As String = "Data Source=ICT-KRITPON\SQLEXPRESS;Initial Catalog=db2012;User ID=sa;Password=$y$05000"
    'Public Shared strConn2 As String = "Data Source=EDP 3\SQLEXPRESS;Initial Catalog=db2012;User ID=sa;Password=$y$05000"
    Public Shared strConn2 As String = "Data Source=192.168.1.5\SQLEXPRESS;Initial Catalog=db2012;User ID=sa;Password=$y$05000"
    'Public Shared strConn2 As String = "Data Source=192.168.1.5\SQLEXPRESS;Initial Catalog=newZone;User ID=sa;Password=$y$05000"
    'Public Shared strConn2 As String = "Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=Test2015;" 'User ID=sa;Password=sys0500"

    '===================================================================================

    'Public Shared strConn2 As String = "Data Source=192.168.1.13\SQLEXPRESS;Initial Catalog=DB2012;User ID=sa;Password=sys0500"
    'Public Shared strConn2 As String = "Data Source=192.168.1.3\SQLEXPRESS;Initial Catalog=DB2012;User ID=sa;Password=$y$05000"

    'strConNet = "server=192.168.1.13\SQLEXPRESS;database=DB2006;Persist Security Info=True;"
    'strConNet = strConNet & "User ID=sa;password=sys0500;"

    Public Shared UserName As String = ""

End Class
