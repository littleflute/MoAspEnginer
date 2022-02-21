function checkstring(name,data,allowednull)
{
var datastr = data;
var lefttrim = datastr.search(/\S/gi);

if (lefttrim == -1) {
    if (allowednull) {
      return 1;
    } else {
      alert("请输入" + name + "！");
      return -2;
    }
  }
  
  if (datastr.search(/[<>]/gi) != -1) {
    alert("" + name + "中包含非法字符<>");
    return -1;
  }
  return 0;
}

function checkemail(name, data, allowednull)
{
  var datastr = data;
  var lefttrim = datastr.search(/\S/gi);
  
  if (lefttrim == -1) {
    if (allowednull) {
      return 1;
    } else {
      alert("请输入一个正确的E-mail地址！");
      return -1;
    }
  }
  var myRegExp = /[a-z0-9](([a-z0-9]|[_\-\.][a-z0-9])*)@([a-z0-9]([a-z0-9]|[_\-][a-z0-9])*)((\.[a-z0-9]([a-z0-9]|[_\-][a-z0-9])*)*)/gi;
  var answerind = datastr.search(myRegExp);
  var answerarr = datastr.match(myRegExp);
  
  if (answerind == 0 && answerarr[0].length == datastr.length)
  {
    return 0;
  }
  
  alert("请输入一个正确的E-mail地址！");
  return -1;
}
function parseNum(theNum){
  if (theNum.substring(0,1)==0)
    theNum=theNum.substring(1)
  return theNum
}
function parseYMD(theYear,theMonth,theDay) {
  theYear=parseNum(theYear)
  theMonth=parseNum(theMonth)
  theDay=parseNum(theDay)
  if ((theYear < 1900) || (theYear > 3000)){
    return 1
  }
  if (theMonth < 1 || theMonth > 12){
    return 2
  }
  if ((theMonth==1 || theMonth==3 || theMonth==5 || theMonth==7 || theMonth==8 || theMonth==10 || theMonth==12) &&
      (theDay <1 || theDay > 31)
     ){
    return 3
  }
  if ((theMonth==4 || theMonth==6 || theMonth==9 || theMonth==11) &&
      (theDay <1 || theDay > 30)
     ){
    return 3
  }
  if (theYear%400==0 || (theYear%4==0 && theYear%100!=0)){  //闰年
    if (theMonth==2 && (theDay <1 || theDay > 29) )
      return 3
  }
  else  //平年
    if (theMonth==2 && (theDay <1 || theDay > 28) )
      return 3
  return 0
}
function isInvalidDate(theDate,separator){
  default_style=1
  if (theDate.length>10 || theDate.length<8)
    return true
  idx1=theDate.indexOf(separator)
  if (idx1==-1)
    return true
  idx2=theDate.indexOf(separator,idx1+1)
  if (idx2==-1)
    return true
  if (isInvalidDate.arguments.length>2)
  	default_style=isInvalidDate.arguments[2]
  if (default_style<1 || default_style>9){
  	alert("传入参数有误！请检查。")
	return true
  }
  if (default_style==1){
  theYear=theDate.substring(0,idx1)
  theMonth=theDate.substring(idx1+1,idx2)
  theDay=theDate.substring(idx2+1)
  }
  if (default_style==2){
  theMonth=theDate.substring(0,idx1)
  theDay=theDate.substring(idx1+1,idx2)
  theYear=theDate.substring(idx2+1)
  }
  if (theDay.length>2)
    return true
  if (parseYMD(theYear,theMonth,theDay)>0)
    return true
  else
    return false
}