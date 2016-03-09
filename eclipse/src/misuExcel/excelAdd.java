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
import org.apache.poi.xslf.model.geom.IfElseExpression;

public class excelAdd {
	private  ArrayList<String> names;
	private  ArrayList<Integer> nones; 
	private  ArrayList<ArrayList<Integer>> addList ;
	
	private Workbook excel,target;
	private String savePath;
	private int sheetNum = 0;
	private int cellNum = 0;
	private int sheetNum_target= 0 ;
	private int cellNum_target = 0;
	private int saveType = 0;
	private String filename ;
	
	public excelAdd(Workbook excel,Workbook target,int sheetNum,int cellNum,int sheetNum_target,
			int cellNum_taget,int saveType,String savePath,String filename){
		this.excel = excel; //母体
		this.target = target; //参考
		this.sheetNum = sheetNum;
		this.cellNum = cellNum;
		this.sheetNum_target = sheetNum_target;
		this.cellNum_target = cellNum_taget;
		this.saveType =saveType;
		this.savePath = savePath;
		this.filename = filename;
	}
	
	private void addTarget(){
		if(target!=null){
		 if(saveType==3){	
Log.info("拼接 目标列");
			Sheet sheet = target.getSheetAt(sheetNum_target);
			names = new ArrayList<String>();
			for(int i=addJpanel.ignore_Rowtar;i<=sheet.getLastRowNum();i++){
				Row row = sheet.getRow(i);
				if(row!=null){
					Cell cell = row.getCell(cellNum_target);
					if(cell!=null){
						names.add(getCellString(cell));
Log.info("目标列增加 "+getCellString(cell));
					}
				}
			}//end for
		 }else if(saveType==4){
Log.info("拼接 目标行");
			 Sheet sheet = target.getSheetAt(sheetNum_target);
				names = new ArrayList<String>();
					Row row = sheet.getRow(cellNum_target);
						if(row!=null){
						 for(int i=addJpanel.ignore_Celltar;i<row.getLastCellNum();i++){
							 Cell cell=row.getCell(i);
							 if(cell!=null){
								names.add(getCellString(cell));
Log.info("目标行增加 "+getCellString(cell));							}
						 }//end for
				}
		 }
		Log.info("names size:"+names.size());
		Log.info("splitTarget is already");
		}
	}
	
	private void examExcel(){
	Log.info("excamExcel");	
		if(excel!=null){
			if(names!=null&&names.size()>0){
				Sheet sheet = excel.getSheetAt(sheetNum);
				initList(names.size());
				nones = new ArrayList<Integer>();
				Boolean isAdd = false;
				for(int j=addJpanel.ignore_Row;j<=sheet.getLastRowNum();j++){
					Row row = sheet.getRow(j);
					if(row!=null){
						Cell cell = row.getCell(cellNum);
						String str = getCellString(cell);
Log.info("对比 " +str);	
						for(int i=0;i<names.size();i++){
							if(str != null&&str.equals(names.get(i))){
								isAdd = true;
								addList.get(i).add(row.getRowNum());
								break;
							}
						}//end names for	
						if(!isAdd){
							nones.add(row.getRowNum());
						}
						isAdd = false;
					}
				}//end for
			}else{
				Log.info("target is none");
			}
		}else{
			Log.info("excel is not exit");
		}
	}
	
	private void examExcel02(){
	Log.info("excamExcel02");	
		if(excel!=null){
			if(names!=null&&names.size()>0){	
				Sheet sheet_add = excel.getSheetAt(sheetNum);
				Sheet sheet = target.getSheetAt(sheetNum_target);
				Row row_add = sheet_add.getRow(cellNum);
				Row row = sheet.getRow(cellNum_target);
				initList(names.size());
				nones = new ArrayList<Integer>();
				Boolean isAdd = false;
					if(row!=null){
						for(int i=addJpanel.ignore_Cell;i<row.getLastCellNum();i++){
							Cell cell = row.getCell(i);
						if(cell!=null){
							String str = getCellString(cell);
Log.info("对比 " +str);	
							for(int k=addJpanel.ignore_Celltar;k<row_add.getLastCellNum();k++){
								Cell cell2 =row_add.getCell(k);
								if(cell2!=null&&str.equals(getCellString(cell2))){
									isAdd = true;		
									addList.get(0).add(k);
									break;
								}
							}
						}
						}//end names for	
						if(!isAdd){
							nones.add(row.getRowNum());
						}
						isAdd = false;
					}
//				}//end for
				Log.info("examExcel is already");
			}else{
				Log.warm("target is none");
			}
		}else{
			Log.warm("excel is not exit");
		}
	}
	
	public void addExcel(){
		if(saveType==3){
			addTarget();
			examExcel();
			new excelWrite(target.getSheetAt(sheetNum_target),excel,saveType,
					savePath,sheetNum,cellNum_target,addList,names,nones,filename);
		}else{
			addTarget();
			examExcel02();
			new excelWrite(target.getSheetAt(sheetNum_target),excel,saveType,
					savePath,sheetNum,cellNum_target,addList,names,nones,filename);
		}
	}
	
	private void initList(int num){
		addList = new ArrayList<ArrayList<Integer>>();
		for(int i=0;i<num;i++){
			addList.add(new ArrayList<Integer>());
		}
	}
	
	private String getCellString(Cell cell){
	  try{
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
	        	 return cell.getRichStringCellValue().getString().trim();
	     }
	  }catch (NullPointerException e) {
			JOptionPane.showMessageDialog(null,e.getMessage(),"错误",JOptionPane.ERROR_MESSAGE);
	  }
		 return null;
	}
}
