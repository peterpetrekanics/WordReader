//This code will read a word doc from C:\ and then convert it to pdf
// Required external jars:
//itextpdf-5.5.4.jar
//poi-3.11-20141221.jar
//poi-ooxml-3.11-20141221.jar
//poi-ooxml-schemas-3.11-20141221.jar
//xmlbeans-2.3.0.jar
// Useful links:
//http://sourceforge.net/projects/itext/files/iText/iText5.5.4/
//http://www.java2s.com/Code/Jar/x/Downloadxmlbeans230jar.htm

import java.io.*;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;

import com.itextpdf.text.pdf.PdfWriter;
import com.itextpdf.text.Document;
import com.itextpdf.text.Paragraph;

public class WordReader {

public static void main(String[] args) {
InputStream fs = null;  
    Document document = new Document();
    XWPFWordExtractor extractor = null ;

try {

    fs = new FileInputStream("C://5612.docx");
    //XWPFDocument hdoc=new XWPFDocument(fs);
    XWPFDocument hdoc=new XWPFDocument(OPCPackage.open(fs));
    //XWPFDocument hdoc=new XWPFDocument(fs);
    extractor = new XWPFWordExtractor(hdoc);
    OutputStream fileOutput = new FileOutputStream(new       File("C://test.pdf"));
    PdfWriter.getInstance(document, fileOutput);
    document.open();
    String fileData=extractor.getText();
    System.out.println(fileData);
    document.add(new Paragraph(fileData));
    System.out.println(" pdf document created");
        } catch(IOException e) {
            System.out.println("IO Exception");
             e.printStackTrace();
          } catch(Exception ex) {
             ex.printStackTrace();
           }finally {  
                document.close();  
           } 
 }//end of main()
}//end of class