package src;
/**
 * Title:	    Export Excel file from Database using XML input.
 * Class:	    ExportExcelV01
 * Description: Export Excel file from Database using XML input.
 * @author  	ArunDavid
 * @version  	1.0
 *
 */
import java.sql.ResultSet;
import java.sql.SQLException;
class KeyMethodMapV1{
	String getMethodValue(String method,ResultSet rs) throws SQLException{
		/*if(method.equalsIgnoreCase("USER_ID_FORMATTED")){
			return getFormattedUserID(rs);
		}
		if(method.equalsIgnoreCase("Txn_Type")){
			return getTxnType(rs);
		}*/
		return null;
	}
	/*String getFormattedUserID(ResultSet rs) throws SQLException{
		return rs.getString("USER_ID");
	}
	String getTxnType(ResultSet rs) throws SQLException{
		return "TYPE:"+rs.getString("TXN_TYPE");
	}*/
}