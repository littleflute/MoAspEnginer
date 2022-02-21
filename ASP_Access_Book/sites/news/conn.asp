<%on error resume next
dim conn
dim connstr
dim db
pathload=""

		dim startime
		dim pass_word,User_ID,Data_Source
		startime=timer()

   	set conn=server.createobject("adodb.connection")
	   connstr="DBQ="+server.mappath("news.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
	conn.Open  connstr

function getchar(strcontent)
changedata=split(strcontent,"<") 
maxtmp=ubound(changedata)
for m=0 to maxtmp
numtm=instr(changedata(m),">")
if numtm<len(changedata(m)) then
contenttmpt=right(changedata(m),len(changedata(m))-numtm)
contenttmp=contenttmp&contenttmpt
end if
next
getchar=replace(contenttmp,"&nbsp;","")
end Function

function gotTopic(str,strlen)  
dim l,t,c  
l=len(str)  
t=0
for k=1 to l
c=Abs(Asc(Mid(str,k,1)))  
if c>255 then
t=t+2 
else
t=t+1
end if
if t>=strlen then
gotTopic=left(str,k)&"..." 
exit for 
else
gotTopic=str
end if
next
end function

function viewclass(classid)
dim cs
set cs=conn.execute("select class from class where classid="&classid)
if not cs.eof then
viewclass=cs(0)
end if
cs.close
end function%>