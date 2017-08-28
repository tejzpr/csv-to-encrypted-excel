import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.*;

class csvtoexcel_threaded {
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
            //FileOutputStream fileOutputStream = new FileOutputStream(xlsxFileAddress);

            final PipedOutputStream fileOutputStream = new PipedOutputStream();
            PipedInputStream in = new PipedInputStream(fileOutputStream);

            new Thread(
                    new Runnable(){
                        public void run(){
                            try {
                                workBook.write(fileOutputStream);
                            } catch (IOException e){

                            }
                        }
                    }
            ).start();

            POIFSFileSystem fs = new POIFSFileSystem();
            EncryptionInfo info = new EncryptionInfo(EncryptionMode.agile);

            Encryptor enc = info.getEncryptor();
            enc.confirmPassword("password123"); // A complex password


            OPCPackage opc = OPCPackage.open(in);
            OutputStream os = enc.getDataStream(fs);
            opc.save(os);
            opc.close();


            FileOutputStream fos = new FileOutputStream(xlsxFileAddress_enc);
            fs.writeFilesystem(fos);
            fos.close();

            System.out.println("Done");

            final long duration = System.nanoTime() - startTime;
            System.out.println("Time to run: "+duration);
        }
        catch ( Exception ex ) {
            System.out.println( ex.getMessage() + "Exception in try" );
        }
    }
}