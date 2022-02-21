<SCRIPT RUNAT=SERVER LANGUAGE=JAVASCRIPT>
//Class Conn()
function Conn() {
	//private variables
	var db = 'data/counter.mdb';
	var connStr = 'Provider = Microsoft.Jet.OLEDB.4.0;Data Source = ' + Server.mapPath(db);
	
	//public variables
	this.conn;
	
	//public function getConn()
	this.getConn = function() {
		try{
			this.conn = Server.createObject('ADODB.Connection');
			this.conn.open(connStr);
		}catch(exception){
			Response.write('<ol>');
			Response.write('<li>' + exception + '</li>');
			Response.write('<li>' + (exception.number & 0xFFFF) + '</li>');
			Response.write('<li>Microsoft Jet connection error</li>');
			//Response.write('<li>' + exception.description + '</li>');		//不打印错误信息，以防止报漏数据库地址
			Response.write('</ol>');
			//throw exception;		//不打印错误信息，以防止报漏数据库地址
			Response.end();			//停止执行，防止下面的错误发生
		}
	}

	//public function execute()
	this.execute = function(sqlStr) {
		try{
			return this.conn.execute(sqlStr);
		}catch(exception){
			Response.write('<ol>');
			Response.write('<li>' + exception + '</li>');
			Response.write('<li>' + (exception.number & 0xFFFF) + '</li>');
			Response.write('<li>' + exception.description + '</li>');
			Response.write('</ol>');
			Response.end();
		}
	}

	//public function close()
	this.close = function() {
		try{
			this.conn.close();
			delete conn;
		}catch(exception){
			Response.write('<ol>');
			Response.write('<li>' + exception + '</li>');
			Response.write('<li>' + (exception.number & 0xFFFF) + '</li>');
			Response.write('<li>' + exception.description + '</li>');
			Response.write('</ol>');
			Response.end();
		}
	}
}
</SCRIPT>