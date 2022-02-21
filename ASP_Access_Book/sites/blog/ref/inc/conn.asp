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
			//Response.write('<li>' + exception.description + '</li>');		//����ӡ������Ϣ���Է�ֹ��©���ݿ��ַ
			Response.write('</ol>');
			//throw exception;		//����ӡ������Ϣ���Է�ֹ��©���ݿ��ַ
			Response.end();			//ִֹͣ�У���ֹ����Ĵ�����
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