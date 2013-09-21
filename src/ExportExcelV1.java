package src;
/**
 * Title:	    Export Excel file from Database using XML input.
 * Class:	    ExportExcelV01
 * Description: Export Excel file from Database using XML input.
 * @author  	ArunDavid
 * @version  	1.0
 *
 */
import java.io.File;
import java.util.List;

import org.jdom.Document;
import org.jdom.Element;
import org.jdom.JDOMException;
import org.jdom.input.SAXBuilder;
 
public class ExportExcelV1 {
	public static void main(String[] args) {
		try {
			SAXBuilder builder = new SAXBuilder();
			File configFile = new File("config.xml");
			Document document = (Document) builder.build(configFile);
			Element rootNode = document.getRootElement();
			List inputs = rootNode.getChild("inputfiles").getChildren("file");
			List constants = rootNode.getChild("constants").getChildren("constant");
			Element style = rootNode.getChild("style");
			ExportTableV1 et=new ExportTableV1(constants);
			Element database=rootNode.getChild("database");
			et.connectDb(database);
			et.setStyle(style);
			for (int i = 0; i < inputs.size(); i++) {
				Element input = (Element) inputs.get(i);
				et.export(input.getText());
			}
			et.closeDb();
		} catch (Exception e) {
			System.out.println(e);
		}
	}
}
