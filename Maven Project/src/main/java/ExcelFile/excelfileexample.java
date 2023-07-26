package ExcelFile;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class excelfileexample {

	XSSFSheet sh;
	public excelfileexample() throws IOException
	{
		FileInputStream f=new FileInputStream("C:\\Users\\Lenovo\\eclipse-workspace\\Maven Project\\src\\main\\resources\\Excelfile.xlsx");
		XSSFWorkbook w=new XSSFWorkbook(f);
		sh=w.getSheet("Sheet1");
	}
	public void readFile()
	{
		for(Row r:sh)
		{
			for(Cell c:r)
			{
				System.out.print(c+ "  ");
			}
			System.out.println();
		}
	}
	public String readData(int i,int j)
	{
		Row r=sh.getRow(i);
		Cell c=r.getCell(j);
	//	Cell c=sh.getRow(i).getCell(j);
		int cellType=c.getCellType();
		switch(cellType)
		{
		case Cell.CELL_TYPE_STRING:
			String data=c.getStringCellValue();
			return data;
		case Cell.CELL_TYPE_NUMERIC:
			double num=c.getNumericCellValue();
			String data2=String.valueOf(num);
			return data2;
			default:
				return null;
			
		}
	}
	public static void main(String[] args) throws IOException {
		excelfileexample ex=new excelfileexample();
		ex.readFile();
		System.out.println("Read data:"+ex.readData(1, 0));
	}

}
