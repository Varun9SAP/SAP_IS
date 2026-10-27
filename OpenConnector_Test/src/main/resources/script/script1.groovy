import com.sap.gateway.ip.core.customdev.util.Message
import org.apache.poi.ss.usermodel.*
import groovy.xml.MarkupBuilder
import java.io.StringWriter

def Message processData(Message message) {
    def body = message.getBody(java.io.InputStream.class)
    Workbook workbook = WorkbookFactory.create(body)
    
    StringWriter writer = new StringWriter()
    MarkupBuilder xml = new MarkupBuilder(writer)
    
    // Wrap everything in a root tag to hold multiple sheets
    xml.WorkbookData {
        // Loop through all available sheets in the file
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i)
            
            // Read the sheet name and remove spaces for valid XML tags
            String sheetName = sheet.getSheetName().replaceAll("\\s+", "_")
            
            // Create a section for each sheet
            xml."${sheetName}" {
                for (Row row : sheet) {
                    xml.Record {
                        // Only check up to the last cell POI detects in this row
                        int lastCell = row.getLastCellNum()
                        
                        for (int cn = 0; cn < lastCell; cn++) {
                            Cell cell = row.getCell(cn, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL)
                            
                            // Skip the cell entirely if it is null or contains only whitespace
                            if (cell != null && cell.toString().trim().length() > 0) {
                                def colName = "Column_" + cn
                                xml."${colName}"(cell.toString().trim())
                            }
                        }
                    }
                }
            }
        }
    }
    
    message.setBody(writer.toString())
    return message
}