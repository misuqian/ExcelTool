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

import java.util.ArrayList;

import javax.swing.JOptionPane;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class excelSplit {
	private static ArrayList<String> names;
	private static ArrayList<Integer> nones; 
	private static ArrayList<ArrayList<Integer>> splitList ;
	
	private Workbook excel,target;
	private String savePath;
	private int sheetNum = 0;
	private int cellNum = 0;
	private int sheetNum_target= 0 ;
	private int cellNum_target = 0;
	private int saveType = 0;
	private String filename;
	
	public excelSplit(Workbook excel,Workbook target,int sheetNum,int cellNum,int sheetNum_target,
			int cellNum_taget,int saveType,String savePath,String filename){
		this.excel = excel;
		this.target = target;
		this.sheetNum = sheetNum;
		this.cellNum = cellNum;
		this.sheetNum_target = sheetNum_target;
		this.cellNum_target = cellNum_taget;
		this.saveType =saveType;
		this.savePath = savePath;
		this.filename = filename;
	}
	
	private void splitTarget(){
		if(target!=null){
			ArrayList<Cell> other = new ArrayList<Cell>();
			Sheet sheet = target.getSheetAt(sheetNum_target);
			names = new ArrayList<String>();
			for(int i=splitJpanel.ignore_Rowtar;i<=sheet.getLastRowNum();i++){
				Row row = sheet.getRow(i);
				if(row!=null){
					Cell cell = row.getCell(cellNum_target);					
					if(cell!=null){
						String str = getCellString(cell);
 					if(!names.contains(str))
 					  if(!other.contains(cell)){
 						 names.add(str);
  						 other.add(cell);
Log.info("分割目标增加 "+str);
 					  }
					}else{
						other.add(cell);
					}
				}
			}//end for
		}
	}
	
	private void examExcel(){
		if(excel!=null){
			if(names!=null&&names.size()>0){
				Sheet sheet = excel.getSheetAt(sheetNum);
				initList(names.size());
				nones = new ArrayList<Integer>();
				Boolean isAdd = false;
				for(int j=splitJpanel.ignore_Row;j<=sheet.getLastRowNum();j++){
					Row row = sheet.getRow(j);
					if(row!=null){
						Cell cell = row.getCell(cellNum);
						String str = getCellString(cell);
Log.info("对比 "+str);
						for(int i=0;i<names.size();i++){
							if(str!=null && str.equals(names.get(i))){
								isAdd = true;
								splitList.get(i).add(row.getRowNum());
							}
						}//end names for	
						if(!isAdd){
							nones.add(row.getRowNum());
						}
						isAdd = false;
					}
				}//end for
				Log.info("examExcel is already");
			}else{
				Log.warm("target is none");
			}
		}else{
			Log.warm("excel is not exit");
		}
	}
	
	public void splitExcel(){
		splitTarget();
		examExcel();
		new excelWrite(excel.getSheetAt(sheetNum),saveType, savePath,splitList,names,nones,filename);
	}
	
	private void initList(int num){
		splitList = new ArrayList<ArrayList<Integer>>();
		for(int i=0;i<num;i++){
			splitList.add(new ArrayList<Integer>());
		}
	}
	
	private String getCellString(Cell cell){
	  try {
		  switch (cell.getCellType()) {
	         case Cell.CELL_TYPE_STRING:
	             return cell.getRichStringCellValue().getString().trim();
	         case Cell.CELL_TYPE_NUMERIC:
	             if (DateUtil.isCellDateFormatted(cell)) {
	            	return cell.getDateCellValue().toString().trim();
	             } else {
	                return String.valueOf(cell.getNumericCellValue()).trim();
	             }
	         case Cell.CELL_TYPE_BOOLEAN:
	            return String.valueOf(cell.getBooleanCellValue()).trim();
	         case Cell.CELL_TYPE_FORMULA:
	        	 return String.valueOf(cell.getCellFormula()).trim();
	         default:
	        	 return cell.getStringCellValue().trim();
	     }
	  } catch (NullPointerException e) {
			JOptionPane.showMessageDialog(null,e.getMessage(),"错误",JOptionPane.ERROR_MESSAGE);
	  }
		 return null;
   }
}
