import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;

public class ExcelHandler {
    public static void main(String[] args) throws IOException, SQLException {
        //连接数据库
        String url = "jdbc:mysql://121.36.55.63:3306/internet?serverTimezone=UTC";
        String username = "Internet";
        String password = "123456";
        Connection conn = DriverManager.getConnection(url, username, password);

        int id = 1;

        File dir = new File("src/excel");
        File[] files = dir.listFiles();

        if(files!=null){
            for(File f:files){
                String path = f.getPath();
                String exRegulationName = f.getName();
                exRegulationName = exRegulationName.substring(0,exRegulationName.lastIndexOf("."));//去掉文件后缀
                InputStream inputStream = Files.newInputStream(Paths.get(path));
                Workbook workbook = new XSSFWorkbook(inputStream);

                Sheet sheet = workbook.getSheetAt(0);
                for(int i=1;i<sheet.getPhysicalNumberOfRows();++i){//遍历2~272行
                    Row row = sheet.getRow(i);
                    Cell flagCell = row.getCell(0);
                    Cell nameCell = row.getCell(4);
                    flagCell.setCellType(CellType.STRING);
                    nameCell.setCellType(CellType.STRING);

                    String flag = flagCell.getStringCellValue();
                    String inRegulationName = nameCell.getStringCellValue();

                    int relative = Integer.parseInt(flag.substring(0,1));
                    if(relative==1){//如果外规与内规相关
                        String sql = "INSERT IGNORE INTO relation VALUES(?,?,?);";
                        PreparedStatement ps = conn.prepareStatement(sql);
                        ps.setInt(1,id);
                        ps.setString(2,exRegulationName);
                        ps.setString(3,inRegulationName);
                        ps.execute();
                        ps.close();
                        id++;
                    }
                }
            }
        }
        conn.close();
    }
}
