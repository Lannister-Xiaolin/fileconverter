package converter;

import com.aspose.words.Document;
import com.aspose.words.FontSettings;
import com.aspose.words.License;
import com.aspose.words.SaveFormat;
import org.apache.commons.lang3.StringUtils;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.tika.metadata.Metadata;
import org.apache.tika.parser.AutoDetectParser;
import org.apache.tika.parser.ParseContext;
import org.apache.tika.sax.BodyContentHandler;

import java.io.*;
import java.util.ArrayList;
import java.util.regex.Matcher;
import java.util.regex.Pattern;


public class FileToText {
    /*
    tika转文本程序，支持格式：docx/doc/pdf(单独使用pdfbox效果更好)/html/xml/txt
    @ Author: Xiaolin
     */
    public static String parse(String file) {
        BodyContentHandler handler = new BodyContentHandler(300000);
        AutoDetectParser parser = new AutoDetectParser();
        Metadata metadata = new Metadata();
        ParseContext parserContext = new ParseContext();
        try (InputStream stream = new FileInputStream(new File(file))) {
            parser.parse(stream, handler, metadata, parserContext);
            return handler.toString();
        } catch (Exception e) {
            e.printStackTrace();
            return "";
        }
    }

    /*
    pdfbox解析pdf，效果优于tika
     */
    public static String pdfParser(String fileString, Boolean textPosition) {
        try {
            File file = new File(fileString);
            PDDocument document = PDDocument.load(file);
            PDFTextStripper pdfStripper = textPosition ? new PDFLayoutTextStripper(): new PDFTextStripper();
            pdfStripper.setSortByPosition(true);
            pdfStripper.setAddMoreFormatting(true);
            String text = pdfStripper.getText(document);
            document.close();
            return text;
        } catch (Exception e) {
            e.printStackTrace();
            return "";
        }
    }
    /*
    pdfbox解析pdf，效果优于tika
     */
    public static String pdfParser(String fileString) {
        try {
            File file = new File(fileString);
            PDDocument document = PDDocument.load(file);
            PDFTextStripper pdfStripper = new PDFTextStripper();
            pdfStripper.setSortByPosition(true);
            pdfStripper.setAddMoreFormatting(true);
            String text = pdfStripper.getText(document);
            document.close();
            return text;
        } catch (Exception e) {
            e.printStackTrace();
            return "";
        }
    }
    /*
    证书有效性验证
    */
    public static boolean getLicense() {
        boolean result = false;
        try {
            InputStream is = Document.class.getResourceAsStream("/com.aspose.words.lic_2999.xml");
            License asposeLic = new License();
            asposeLic.setLicense(is);
            result = true;
            is.close();
            return result;
        } catch (Exception e) {
            e.printStackTrace();
            return result;
        }
    }

    /*
    convert doc or docx to pdf
    */
    public static void docToPdf(String inPath, String outPath) throws Exception {
        if (!getLicense()) { // match License if not marks on document
            throw new Exception("license not correct!");
        }
        System.out.println(inPath + " -> " + outPath);
        try {
            File file = new File(outPath);
            FileOutputStream os = new FileOutputStream(file);
            Document doc = new Document(inPath); // doc/docx
            FontSettings.getDefaultInstance().setFontsFolders(new String[]{"/usr/share/fonts/truetype/chinese", "C:\\Windows\\Fonts"}, true);
            doc.save(os, SaveFormat.PDF);
            os.close();  //关闭文件
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /*
    aspose convert doc to text
     */
    public static void docToText(String inPath, String outPath) throws Exception {
        if (!getLicense()) { // match License if not marks on document
            throw new Exception("license not correct!");
        }
        System.out.println(inPath + " -> " + outPath);
        try {
            long old = System.currentTimeMillis();
            File file = new File(outPath);
            FileOutputStream os = new FileOutputStream(file);
            Document doc = new Document(inPath); // doc/docx
            FontSettings.getDefaultInstance().setFontsFolders(new String[]{"/usr/share/fonts/truetype/chinese", "C:\\Windows\\Fonts"}, true);
            doc.save(os, SaveFormat.TEXT);
            long now = System.currentTimeMillis();
            System.out.println("convert OK! " + ((now - old) / 1000.0) + "秒");
            os.close();  //关闭文件
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void writeFile(String fileName, String text) {
        try {
            File writeName = new File(fileName); // 相对路径，如果没有则要建立一个新的output.txt文件
            writeName.createNewFile(); // 创建新文件,有同名的文件的话直接覆盖
            try (FileWriter writer = new FileWriter(writeName); BufferedWriter out = new BufferedWriter(writer)) {
                String[] textArrayStrings = text.split("\n|\r");
                text = "";
//	            	System.out.print(textArrayStrings.length);
                for (int i = 0; i < textArrayStrings.length; i++) {
                    if (textArrayStrings[i].isEmpty()) {
                        continue;
                    }
                    text += StringUtils.strip(textArrayStrings[i]) + "\n";
                }
                out.write(text); // \r\n即为换行
                out.flush(); // 把缓存区内容压入文件
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static ArrayList<String> getFiles(String path) {
        ArrayList<String> files = new ArrayList<String>();
        File file = new File(path);
        File[] tempList = file.listFiles();
        String pattenString = ".doc$|.docx$|.pdf$|html$";
        Pattern pattern = Pattern.compile(pattenString);
        for (File value : tempList) {
            if (value.isFile()) {
                Matcher matcher = pattern.matcher(value.toString());
                if (matcher.find())
                    System.out.println("文     件：" + value);
                files.add(value.toString());
            }
            if (value.isDirectory()) {
                System.out.println("文件夹：" + value);
            }
        }
        return files;
    }

    public static void bathConvert(String path) {
        ArrayList<String> files = getFiles(path);
//		    String[] fileListString = files.toArray();
        String resulString = "";
        String txtFileName = "";
        for (String file : files) {
            long startTime = System.currentTimeMillis();
            String fileString = file;
            try {
                resulString = parse(fileString);
            } catch (Exception e) {
                // TODO Auto-generated catch block
                e.printStackTrace();
            }
            String pattenString = ".doc$|.docx$|.pdf$|.html$";
            Pattern pattern = Pattern.compile(pattenString);
            Matcher matcher = pattern.matcher(fileString);
            txtFileName = matcher.replaceFirst("_javainterface.txt");
            writeFile(txtFileName, resulString);
            long endTime = System.currentTimeMillis(); // 获取结束时间
            System.out.println("程序运行时间： " + (endTime - startTime) + "ms");
        }
    }

    public static void main(String[] args) {
        try {
            long startTime = System.currentTimeMillis();
//            System.out.println(pdfParser("F:\\Download\\【.Net开发工程师（2019届）_成都】游翼嘉 应届生 - 副本.pdf", true));
            docToPdf("E:\\Programming\\Python\\4_NLP\\resume_analysis\\数据集\\+解析样本记录\\【.Net系统架构师 _ 广州 20k-30k】侯宜安 15年.docx",
                    "E:\\【.Net系统架构师 _ 广州 20k-30k】侯宜安 15年.pdf");
            long endTime = System.currentTimeMillis(); // 获取结束时间
            System.out.println("程序运行时间： " + (endTime - startTime) + "ms");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
