<%
foundUser = FALSE
Application.Lock
localStats = Application( "Stats" )
FOR k = 0 TO UBOUND(localStats, 2)
  IF localStats( 0, k ) = Session.SessionID THEN
    localStats( 1, k ) =  Request.ServerVariables( "SCRIPT_NAME" )
    foundUser = TRUE
    EXIT FOR
  END IF
NEXT
IF foundUser = FALSE THEN
FOR k = 0 TO UBOUND( localStats, 2 )
  IF localStats( 0, k ) = "" THEN
    localStats( 0, k ) = Session.SessionID
    localStats( 1, k ) = Request.ServerVariables( "SCRIPT_NAME" )
    localStats( 2, k ) = Request.ServerVariables( "REMOTE_ADDR" )
    localStats( 3, k ) = NOW()  
    EXIT FOR
  END IF
NEXT
END IF
Application("Stats") = localStats
Application.UnLock
%>
