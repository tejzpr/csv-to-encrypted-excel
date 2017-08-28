import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.*;

public class csvtoexcel_buffer {
    public static void main(String args[]) {
        try {
            final long startTime = System.nanoTime();
            //   String fName = args[ 0 ];

            String csvFileAddress = "./csv/AllstarFull.csv"; //csv file address
            String xlsxFileAddress_enc = "./excel/AllstarFull.xlsx"; //xlsx file address

            final SXSSFWorkbook workBook = new SXSSFWorkbook();
            Sheet sheet = workBook.createSheet("sheet1");
            String currentLine = null;
            int RowNum = 0;
            BufferedReader br = new BufferedReader(new FileReader(csvFileAddress));
            while ((currentLine = br.readLine()) != null) {
                String str[] = currentLine.split(",");
                RowNum++;
                Row currentRow = sheet.createRow(RowNum);
                for (int i = 0; i < str.length; i++) {
                    currentRow.createCell(i).setCellValue(str[i]);
                }
            }


            POIFSFileSystem fs = new POIFSFileSystem();
            EncryptionInfo info = new EncryptionInfo(EncryptionMode.agile);

            Encryptor enc = info.getEncryptor();
            enc.confirmPassword("password123"); // A complex password

            ByteArrayOutputStream baos = null;
            ByteArrayInputStream bais = null;

            FileOutputStream fos = new FileOutputStream(xlsxFileAddress_enc);

            try {
                baos = new ByteArrayOutputStream();
                workBook.write(baos);
                bais = new ByteArrayInputStream(baos.toByteArray());

                OPCPackage opc = OPCPackage.open(bais);
                OutputStream os = enc.getDataStream(fs);
                opc.save(os);
                opc.close();
            }
            catch (Exception e) {
                throw new IllegalStateException("Error writing encrypted Excel document", e);
            }
            finally {
                IOUtils.closeQuietly(baos);
                IOUtils.closeQuietly(bais);
            }

            fs.writeFilesystem(fos);

            fos.close();
            final long duration = System.nanoTime() - startTime;
            System.out.println("Time to run: "+duration);
        }
        catch (IOException e) {
        }
    }
}
