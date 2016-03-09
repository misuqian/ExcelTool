/*
 Copyright (c) 2016 pengqian
 
 Permission is hereby granted, free of charge, to any person obtaining a copy
 of this software and associated documentation files (the "Software"), to deal
 in the Software without restriction, including without limitation the rights
 to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 copies of the Software, and to permit persons to whom the Software is
 furnished to do so, subject to the following conditions:
 
 The above copyright notice and this permission notice shall be included in all
 copies or substantial portions of the Software.
 
 THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 SOFTWARE.
 */


package misuExcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class excelRead {
	private String pathString = null;	
	private ArrayList<String> listPath =null;
	private int type = 0;
	private  Workbook wb;
//	public static excelRead excel;
	
	public excelRead(ArrayList<String> listPath,int type){
		this.type = type;
		this.listPath = listPath;
//		setPath(pathString);
//		createExcel();
	}
	
	
	public excelRead(String pathString,int type){
		this.type = type;
		setPath(pathString);
		createExcel();
	}
	
	public void setPath(String str){
		this.pathString= str;
	}
	
	public File getFile(){ 	
		return (pathString==null)? null: new File(pathString);
	}
	
	public void createExcel(){
		InputStream inps = null ; 
		try {
			inps = new FileInputStream(getFile());
		} catch (FileNotFoundException e) {
			Log.warm(e.getCause().getMessage());
		}
		
		if(inps!=null){
			try {
				wb=WorkbookFactory.create(inps);
				switch (type) {
				case 1:
//					windows.readyEx01=true;
//					System.err.println("first excel is already");
					break;
				case 2:
//					System.err.println("second already");
//					windows.readyEx02=true;
					break;

				default:
					break;
				}
			} catch (InvalidFormatException e) {
				Log.warm(e.getCause().getMessage());
			} catch (IOException e) {
				Log.warm(e.getCause().getMessage());
			}
		}
	}
	
	public Workbook getWorkbook(){
		return wb;
	}
	
	public int getSheetCount(){
		return wb.getNumberOfSheets();
	}
	
	public ArrayList<String> getSheetNames(){
		ArrayList<String> names = new ArrayList<String>();
		for(int i=0;i<wb.getNumberOfSheets();i++){
			names.add(wb.getSheetName(i));
		}
		return names;
	}
	
	public int getSCellNum(int i){
		if(i>-1){
			Sheet sheet = (Sheet) wb.getSheetAt(i);
			int max = 0;
			for(int j =0;j<sheet.getLastRowNum();j++){
				Row row = sheet.getRow(j);
				int r = row.getLastCellNum();
				if(r>max){
					max = r;
				}
			}
			return max;
		}
	 return 0;
	}
	
	public int getSRowNum(int i){
		 if(i>-1){
			 Sheet sheet = (Sheet) wb.getSheetAt(i);
			 if(sheet!=null){
				 return (sheet.getLastRowNum()+1);
			 }
		}
			 return 0;
		}
	
}
