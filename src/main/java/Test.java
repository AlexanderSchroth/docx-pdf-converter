

import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

public class Test {

    public static void main(String[] args) throws Exception {
        InputStream doc = new FileInputStream(new File("./test.docx"));
        XWPFDocument document = new XWPFDocument(doc);
        PdfOptions options = PdfOptions.create();

        OutputStream out = new FileOutputStream(new File("./test.pdf"));
        PdfConverter.getInstance().convert(document, out, options);
        System.out.println("Done");
    }
}
