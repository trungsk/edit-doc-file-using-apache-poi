import com.ibm.icu.text.NumberFormat;
import com.ibm.icu.text.RuleBasedNumberFormat;
import org.apache.commons.io.FileUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.util.*;


/** the Flow:
 *  the CCQMContractTemplate.docx has a lot of blanks for infomations. a unique variable is put on each blank (we set it in camel case)
 *  we are going to create a map with its keys named after the variables in the docx and its values are what will be filled in the blank
 *  After programme runs, there are 3 files generated in folder template (write some codes to auto-delete them if you want to or do it manually)
 *          1, contract.docx is our contract after filled by the values in the map
 *          2, base64.txt is our text file which contains base64-code encoded from contract.docx
 *          3, encoded-contract.docx is the doc file decoded from base64-code comes from base64.txt file
 *          4, actually there is one more file as a clone file and actually the contract.docx are going to be generated from this clone
 *              not from the template. Because our method would overwrite the values into the template so a clone have to be created for
 *              sort of sacrificing. It will be created then deleted in the blink of an eye that we can not see its existence in the project's tree.
 *              I know there are many samples that can manage the project without generating a clone file but I find it difficult
 *              and complicated for newbies. Everything will be done with this short-and-sweet method with a tiny clone file.
 *
 *  The values we'll use here are static not from a database. Create one and do custom your own values in value-part of the map below as your requests
 *  
 *
 */


public class Main {
    public static void main(String[] args) {

        Map<String, Object> map = new HashMap<>();
        map.put("currentDate", Calendar.getInstance().get(Calendar.DATE));
        map.put("currentMonth", Calendar.getInstance().get(Calendar.MONTH) + 1); // add 1 because month count from 0
        map.put("currentYear", Calendar.getInstance().get(Calendar.YEAR));
        map.put("bankBranch", "HOAN KIEM");
        map.put("bankAddress", "25 Tran Hung Dao");
        map.put("bankNumber", "1900 555 587");
        map.put("bankFax", "");
        map.put("representer", "PHAM DINH CUONG");
        map.put("position", "Expert");
        map.put("cusName", "NGUYEN THANH TRUNG");
        map.put("cusIDNumber", "1597858955");
        map.put("cusIDIssuedOn", "23-4-2005");
        map.put("cusIDFrom", "CA HA NOI");
        map.put("cusAddress", "Số 2 Tràng Thi");
        map.put("cusWorkPlace", "47 Phạm Văn Đồng");
        map.put("cusAccountNumber", "19055632485715");
        map.put("cusAccountBranch", "HOANG MAI");
        map.put("cusPhoneNumber", "02568965574");
        map.put("serviceFee", "23500007804");
        map.put("moneyText", convertMoneyToText(map.get("serviceFee").toString()));
        String source = "template/CCQMContractTemplate.docx";
        String copy = "template/temp.docx";
        String des = "template/contract.docx";
        File fileSource = new File(source);
        File fileCopy = new File(copy);
        try {
            copyFileUsingStream(fileSource, fileCopy);
            updateDoc(copy, des, map);
            // write String to txt file using apache common io
            String base64 = encodeToBase64(des);
            FileUtils.writeStringToFile(new File("template/base64.txt"), base64);
            // decode base64 to docx
            byte[] decodedBytes = Base64.getDecoder().decode(base64);
            // Note the try-with-resources block here, to close the stream automatically
            try (OutputStream stream = new FileOutputStream("template/encoded-contract.docx")) {
                stream.write(decodedBytes);
            }
            System.out.println("your contract has been created!");
        } catch (IOException e) {
            System.out.println("your contract has not been created! There must be some errors");
            e.printStackTrace();
        } finally {
            fileCopy.delete();
        }


    }
    private static void updateDoc(String input, String output, Map<String, Object> map) throws IOException {
        try (XWPFDocument doc = new XWPFDocument(OPCPackage.open(input))) {
            for (XWPFParagraph p : doc.getParagraphs()) {
                List<XWPFRun> runs = p.getRuns();
                if (runs != null) {
                    for (XWPFRun r : runs) {
                        String text = r.getText(0);
                        for (Map.Entry<String, Object> entry : map.entrySet()) {
                            if (text != null && text.contains(entry.getKey())) {
                                text = text.replace(entry.getKey(), entry.getValue().toString());
                                r.setText(text, 0);
                            }
                        }
                    }
                }
            }
            for (XWPFTable tbl : doc.getTables()) {
                for (XWPFTableRow row : tbl.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        for (XWPFParagraph p : cell.getParagraphs()) {
                            for (XWPFRun r : p.getRuns()) {
                                String text = r.getText(0);
                                for (Map.Entry<String, Object> entry : map.entrySet()) {
                                    if (text != null && text.contains(entry.getKey())) {
                                        text = text.replace(entry.getKey(), entry.getValue().toString());
                                        r.setText(text, 0);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            doc.write(new FileOutputStream(output));
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }

    }

    private static void copyFileUsingStream(File source, File dest) throws IOException {
        InputStream is = null;
        OutputStream os = null;
        try {
            is = new FileInputStream(source);
            os = new FileOutputStream(dest);
            byte[] buffer = new byte[1024];
            int length;
            while ((length = is.read(buffer)) > 0) {
                os.write(buffer, 0, length);
            }
        } finally {
            is.close();
            os.close();
        }
    }

    private static String encodeToBase64(String pathname) {
        File originalFile = new File(pathname);
        String encodedBase64 = null;
        try {
            FileInputStream fileInputStreamReader = new FileInputStream(originalFile);
            byte[] bytes = new byte[(int) originalFile.length()];
            fileInputStreamReader.read(bytes);
            encodedBase64 = Base64.getEncoder().encodeToString(bytes);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        return encodedBase64;
    }

    public static String convertMoneyToText(String input){
        String output = "";
        try {
            NumberFormat ruleBasedNumberFormat = new RuleBasedNumberFormat(new Locale("vi","VN"), RuleBasedNumberFormat.SPELLOUT);
            output = ruleBasedNumberFormat.format(Long.parseLong(input)) + " đồng";
        }catch (Exception e){
            output = "không đồng";
        }
        return output;
    }


}

