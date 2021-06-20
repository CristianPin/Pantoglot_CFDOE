package datos;

import com.spire.doc.Document;
import com.spire.doc.TextWatermark;
import com.spire.doc.documents.WatermarkLayout;
import java.awt.Color;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class ControlArchivos {

    public File NewRuta;
    File F1, F2, F3, FF;
    public static String District, County, ContactFirstName, ContactLastName,
            Salutation, ProfessionalTitle, JobTitle1, JobTitle2, Phone, Email,
            Address, City, State, ZIPCode, FinalRut, RutaRuta, Test, RutaTest, FPrincipal, FTres, FDos, RutaCompleta, nombreArchivo;

    public void LeerEXCEL() {
        List cellData = new ArrayList();
        try {
            FileInputStream file = new FileInputStream(new File(nombreArchivo));
            XSSFWorkbook libro = new XSSFWorkbook(file);
            XSSFSheet sheet = libro.getSheetAt(0);
            Iterator rowIterator = sheet.rowIterator();

            while (rowIterator.hasNext()) {
                XSSFRow row = (XSSFRow) rowIterator.next();
                Iterator iterator = row.cellIterator();
                List cellTempo = new ArrayList();
                while (iterator.hasNext()) {
                    XSSFCell cell = (XSSFCell) iterator.next();
                    cellTempo.add(cell);
                }
                cellData.add(cellTempo);
            }
        } catch (Exception e) {
            System.out.println("Error Control.LeerEXCEL: "+e);
        }

        obtenerDatosExcel(cellData);
    }

    private static void obtenerDatosExcel(List cellDataList) {
        for (int i = 0; i < cellDataList.size(); i++) {
            List cellTempList = (List) cellDataList.get(i);
            for (int j = 0; j < cellTempList.size(); j++) {

                XSSFCell Cell = (XSSFCell) cellTempList.get(j);
                String CellValue = Cell.toString();

                XSSFCell CellValidate = (XSSFCell) cellTempList.get(0);
                String Validate = CellValidate.toString();
                //District,County,ContactFirstName,ContactLastName,
                //Salutation,ProfessionalTitle,JobTitle1,JobTitle2,Phone,Email,Address,City,State,ZIPCode;
                if (Validate.equals("1.0")) {

                    switch (j) {
                        case 0:
                            System.out.println("");
                            break;
                        case 1:
                            District = CellValue;
                            break;
                        case 2:
                            County = CellValue;
                            break;
                        case 3:
                            ContactFirstName = CellValue;
                            break;
                        case 4:
                            ContactLastName = CellValue;
                            break;
                        case 5:
                            Salutation = CellValue;
                            break;
                        case 6:
                            ProfessionalTitle = CellValue;
                            break;
                        case 7:
                            JobTitle1 = CellValue;
                            break;
                        case 8:
                            JobTitle2 = CellValue;
                            break;
                        case 9:
                            Phone = CellValue;
                            break;
                        case 10:
                            Email = CellValue;
                            break;
                        case 11:
                            Address = CellValue;
                            break;
                        case 12:
                            City = CellValue;
                            break;
                        case 13:
                            State = CellValue;
                            break;
                        case 14:
                            ZIPCode = CellValue.substring(0, CellValue.length() - 2);
                            break;
                    }
                }

            }

        }
    }

    public void Carpetas(File Ruta, int Validar) {
        File[] Archivos = Ruta.listFiles();
        String[] Carpetas = Ruta.list();

        if (Archivos == null) {
            System.out.println("No hay ficheros en el directorio especificado");
        } else {
            for (int x = 0; x < Archivos.length; x++) {
                if (Archivos[x].isDirectory()) {
                    //System.out.println(Archivos[x]);
                    //System.out.println(Carpetas[x]);
                    switch (Validar) {
                        case 1:
                            FPrincipal = Carpetas[x];
                            RutaCompleta = NewRuta + "\\" + FPrincipal;
                            System.out.println(RutaCompleta);
                            F1 = new File(NewRuta + "\\" + FPrincipal);
                            F1.mkdir();
                            Carpetas(Archivos[x], 2);
                            break;
                        case 2:
                            if (x - 1 != Archivos.length) {
                                FDos = FPrincipal + "\\" + Carpetas[x];
                                RutaCompleta = NewRuta + "\\" + FDos;
                                System.out.println(RutaCompleta);
                                F2 = new File(NewRuta + "\\" + FDos);
                                F2.mkdir();
                                Carpetas(Archivos[x], 3);
                            }
                            break;
                        case 3:
                            if (x - 1 != Archivos.length) {
                                FTres = FDos + "\\" + Carpetas[x];
                                RutaCompleta = NewRuta + "\\" + FTres;
                                System.out.println(RutaCompleta);
                                F3 = new File(NewRuta + "\\" + FTres);
                                F3.mkdir();
                                Carpetas(Archivos[x], 4);
                            }
                            break;
                    }
                    
                } else {
                    System.out.println(Carpetas[x]);
                    File Archivo = new File(Ruta + "\\" + Carpetas[x]);
                    //System.out.println(Archivo);
                    String Name = Carpetas[x];
                    String Ncadena = Name.substring(0, Name.length() - 5);
                    System.out.println(Ncadena);
                    File FF = new File(RutaCompleta);
                    BuscarReemplazarFooterHeader(Archivo, FF, Ncadena);
                }
            }
        }
    }

    public static void BuscarReemplazarFooterHeader(File Archivo, File Ruta, String Ficheros) {
        
        try {
            XWPFDocument doc = new XWPFDocument(new FileInputStream(new File(Archivo.toString())));
            XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(doc);
            
            List<XWPFHeader> headers = doc.getHeaderList();
            for (XWPFHeader h : headers) {
                for (XWPFTable tbl : h.getTables()) {
                    for (XWPFTableRow row : tbl.getRows()) {
                        for (XWPFTableCell cell : row.getTableCells()) {
                            for (XWPFParagraph p : cell.getParagraphs()) {
                                for (XWPFRun r : p.getRuns()) {
                                    String text = r.getText(0);
                                    //System.out.println(text);

                                    switch (text) {
                                        case "Salutation ":
                                            text = text.replace("Salutation ", Salutation + " ");
                                            r.setText(text, 0);
                                            break;
                                        case "ContactFirstName":
                                            text = text.replace("ContactFirstName", ContactFirstName);
                                            r.setText(text, 0);
                                            break;
                                        case "ContactLastName":
                                            text = text.replace("ContactLastName", ContactLastName);
                                            r.setText(text, 0);
                                            break;
                                        case "ProfessionalTitle":
                                            text = text.replace("ProfessionalTitle", ProfessionalTitle);
                                            r.setText(text, 0);
                                            break;
                                        case "JobTitle1":
                                            text = text.replace("JobTitle1", JobTitle1);
                                            r.setText(text, 0);
                                            break;
                                        case "JobTitle2":
                                            text = text.replace("JobTitle2", JobTitle2);
                                            r.setText(text, 0);
                                            break;
                                        case "Phone":
                                            text = text.replace("Phone", Phone);
                                            r.setText(text, 0);
                                            break;
                                        case "E-mail":
                                            text = text.replace("E-mail", Email);
                                            r.setText(text, 0);
                                            break;
                                        case "District":
                                            text = text.replace("District", District);
                                            r.setText(text, 0);
                                            break;
                                        case "County":
                                            text = text.replace("County", County);
                                            r.setText(text, 0);
                                            break;
                                        case "State":
                                            text = text.replace("State", State);
                                            r.setText(text, 0);
                                            break;
                                        case "Address":
                                            text = text.replace("Address", Address);
                                            r.setText(text, 0);
                                            break;
                                        case "City":
                                            text = text.replace("City", City);
                                            r.setText(text, 0);
                                            break;
                                        case "ZipCode":
                                            text = text.replace("ZipCode", ZIPCode);
                                            r.setText(text, 0);
                                            break;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            List<XWPFFooter> Footers = doc.getFooterList();
            for (XWPFFooter F : Footers) {
                for (XWPFTable tbl : F.getTables()) {
                    for (XWPFTableRow row : tbl.getRows()) {

                        for (XWPFTableCell cell : row.getTableCells()) {
                            for (XWPFParagraph p : cell.getParagraphs()) {
                                for (XWPFRun r : p.getRuns()) {
                                    String text = r.getText(0);
                                    //System.out.println(text);

                                    if (text != null) {
                                        switch (text) {
                                            case "District":

                                                text = text.replace("District", District);
                                                r.setText(text, 0);
                                                break;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            //doc.enforceReadonlyProtection(); //enforce readonly protection
            FileOutputStream out = new FileOutputStream(Ruta.toString() + "\\" + Ficheros + "_CF.docx");
            doc.write(out);
            out.close();
            out.close();

            String RutaFinish = Ruta.toString() + "\\" + Ficheros + "_CF.docx";
            AbrirWordwithSpire(RutaFinish);
            // Codigo para crear en PDF
            /*FinalRut = Ruta.toString() + "\\" + Ficheros + ".pdf";
            RutaRuta = FinalRut + ".pdf";
            //Convierte POI XWPFDocument a PDF
            File outFile = new File(Ruta.toString() + "\\" + Ficheros + ".pdf");
            
            //outFile.getParentFile().mkdirs();
            OutputStream out = new FileOutputStream(outFile);
            PdfOptions options = null;//PdfOptions.create();//.fontEncoding("ISO-8859-1");
            PdfConverter.getInstance().convert(doc, out, options);*/
        } catch (Exception e) {
            System.out.println("Error Control.BuscarReemplazarFooterHeader: "+e);
        }

    }

    private static void AbrirWordwithSpire(String Ruta) {
        Document document = new Document();
        document.loadFromFile(Ruta);
        InsertarMarcaAguaSpire(document);
        document.saveToFile(Ruta);
        document.close();
    }

    private static void InsertarMarcaAguaSpire(Document document) {
        TextWatermark txtWatermark = new TextWatermark();
        txtWatermark.setText(District);
        txtWatermark.setFontSize(20);
        txtWatermark.setColor(Color.black);
        txtWatermark.setLayout(WatermarkLayout.Diagonal);
        document.setWatermark(txtWatermark);
    }

}
