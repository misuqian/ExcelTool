package misuExcel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Properties;

import javax.swing.JOptionPane;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.functions.Index;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class excelWrite {
	private Sheet wbSheet = null;
	private Workbook addWb = null;
	private ArrayList<String> names = null;
	private ArrayList<Integer> nones = null; 
	private ArrayList<ArrayList<Integer>> list =null;
	private int type = 1; 
	private int prolevel = 1;
	private String savePath;
	private int sheetNum_target;
	private int cellNum_target;
	private String filename;
	private String fileReal;
	private String _index =".xlsx";
	private int indexType = 1;
	
	public excelWrite(Sheet wbSheet,int type,String savePath,
			ArrayList<ArrayList<Integer>> list,ArrayList<String> names,ArrayList<Integer> nones,String filename){
		this.wbSheet = wbSheet;
		this.type = type;
		this.savePath=savePath;
		this.list =	list;
		this.names= names;
		this.nones= nones;
		this.filename = filename;
		this.fileReal = getRealName(filename);
		output();
	}
	
	private String getRealName(String str){
	System.err.println(str);
		String[] strs = str.split("\\.");
		if(strs.length!=0){
			_index  = strs[strs.length-1];
			if(_index.equals("xls")){
				indexType = 2;
			}else if(_index.equals("xlsx")||_index.equals("")){
				indexType = 1;
			}
			return strs[0];
		}
		return str;
	}
	
	public excelWrite(Sheet wbSheet,Workbook addWb,int type,String savePath,int sheetNum_target,int cellNum_target,
			ArrayList<ArrayList<Integer>> list,ArrayList<String> names,ArrayList<Integer> nones,String filename){
		this.wbSheet = wbSheet;
		this.addWb = addWb;
		this.type = type;
		this.savePath=savePath;
		this.cellNum_target = cellNum_target;
		this.sheetNum_target = sheetNum_target;
		this.list =	list;
		this.names= names;
		this.nones= nones;
		this.filename = filename;
		this.fileReal = getRealName(filename);
		output();
	}
	
	private void output(){
	Log.info("output_type:"+type);
	createFolder();
		switch (type) {
		case 1:
			outType01();
			break;
		case 2:
			outType02();
			break;
		case 3:
			outType03();
			break;
		case 4:
			outType04();
			break;
		default:
			break;
		}
	}
	
	private void outType01(){
	 if(wbSheet!=null&&names!=null&&list!=null){
Log.info("list size:"+list.size());	
		 String strinfo ="";
			for(int i=0;i<list.size();i++){
				ArrayList<Integer> integers = list.get(i);
				Workbook splitWb = null;
			if(indexType==1)
				splitWb=new XSSFWorkbook();
			else if(indexType==2)
				splitWb=new HSSFWorkbook();
				Sheet sheet = splitWb.createSheet("split");
				for(int j=0;j<integers.size()+splitJpanel.ignore_Row;j++){
					Row row=null;
					Row copy= null;						
				 if(j>=splitJpanel.ignore_Row){
					 row =sheet.createRow(j);
					 copy= wbSheet.getRow(integers.get(j-splitJpanel.ignore_Row));	 
				 }else{
					 row =sheet.createRow(j);
					 copy= wbSheet.getRow(j);
				 }
					for(int k=0;k<copy.getLastCellNum();k++){
						Cell cell =row.createCell(k);
						Cell copyCell = copy.getCell(k);
					if(copyCell!=null){	
					    switch (copyCell.getCellType()) {
			                case Cell.CELL_TYPE_STRING:
			                	cell.setCellValue(copyCell.getRichStringCellValue().getString().trim());
			                    break;
			                case Cell.CELL_TYPE_NUMERIC:
			                    if (DateUtil.isCellDateFormatted(copyCell)) {
			                    	cell.setCellValue(copyCell.getDateCellValue());
			                    } else {
			                    	cell.setCellValue(copyCell.getNumericCellValue());
			                    }
			                    break;
			                case Cell.CELL_TYPE_BOOLEAN:
			                	cell.setCellValue(copyCell.getBooleanCellValue());
			                    break;
			                case Cell.CELL_TYPE_FORMULA:
			                	cell.setCellValue(copyCell.getCellFormula());
			                    break;
			                default :
			                	cell.setCellValue(copyCell.getStringCellValue().trim());
					    }
					}  
				  }
				}
				createWB(splitWb,names.get(i));
Log.info(names.get(i)+".xlsx完成");		
				strinfo +=names.get(i)+"."+_index+"完成;";
				if(i!=0 && i%3==0){
					strinfo += "\n";
				}
			}//end for
			JOptionPane.showMessageDialog(null,strinfo);
		}
	}
	
	private void outType02(){
		 if(wbSheet!=null&&names!=null&&list!=null){
Log.info("list size:"+list.size());	
				Workbook splitWb = null;
				if(indexType==1)
					splitWb=new XSSFWorkbook();
				else if(indexType==2)
					splitWb=new HSSFWorkbook();
					for(int i=0;i<list.size();i++){
						ArrayList<Integer> integers = list.get(i);
						Sheet sheet = splitWb.createSheet(names.get(i));
						for(int j=0;j<integers.size()+splitJpanel.ignore_Row;j++){
							Row row=null;
							Row copy= null;
						 if(j>=splitJpanel.ignore_Row){
							 row =sheet.createRow(j);
							 copy= wbSheet.getRow(integers.get(j-splitJpanel.ignore_Row));
						 }else{
							 row =sheet.createRow(j);
							 copy= wbSheet.getRow(j);
						 }
							for(int k=0;k<copy.getLastCellNum();k++){
								Cell cell =row.createCell(k);
								Cell copyCell = copy.getCell(k);
							if(copyCell!=null){	
							    switch (copyCell.getCellType()) {
					                case Cell.CELL_TYPE_STRING:
					                	cell.setCellValue(copyCell.getRichStringCellValue().getString().trim());
					                    break;
					                case Cell.CELL_TYPE_NUMERIC:
					                    if (DateUtil.isCellDateFormatted(copyCell)) {
					                    	cell.setCellValue(copyCell.getDateCellValue());
					                    } else {
					                    	cell.setCellValue(copyCell.getNumericCellValue());
					                    }
					                    break;
					                case Cell.CELL_TYPE_BOOLEAN:
					                	cell.setCellValue(copyCell.getBooleanCellValue());
					                    break;
					                case Cell.CELL_TYPE_FORMULA:
					                	cell.setCellValue(copyCell.getCellFormula());
					                    break;
					                default :
					                	cell.setCellValue(copyCell.getStringCellValue().trim());
							    }
							}  
						  }
						}
					}//end for
					createWB(splitWb,fileReal+"(cut)");
					JOptionPane.showMessageDialog(null,fileReal+"(cut)."+_index+"完成");
				}
	}
	
	private void outType03(){
		 if(wbSheet!=null&&addWb!=null&&names!=null&&list!=null){
			 		Sheet sheet = addWb.getSheetAt(sheetNum_target);
				for(int i=0;i<list.size();i++){
						ArrayList<Integer> integers = list.get(i);
						Row copy= wbSheet.getRow(i+addJpanel.ignore_Rowtar);
						for(int j=0;j<integers.size();j++){
							Row row =sheet.getRow(integers.get(j));	
							int numRow = row.getLastCellNum();	
							for(int k=addJpanel.ignore_Celltar;k<copy.getLastCellNum();k++){
								Cell cell =null;
								Cell copyCell=null;
							if(k!=cellNum_target){
								copyCell = copy.getCell(k);
								if(addJpanel.ignore_Celltar>cellNum_target){
									cell =row.createCell(k+numRow-addJpanel.ignore_Celltar);
								}else {
									cell =row.createCell(k<cellNum_target?(k+numRow-addJpanel.ignore_Celltar):(k-1+numRow-addJpanel.ignore_Celltar));
								}	
							}
							if(copyCell!=null){	
							    switch (copyCell.getCellType()) {
					                case Cell.CELL_TYPE_STRING:
					                	cell.setCellValue(copyCell.getRichStringCellValue().getString().trim());
					                    break;
					                case Cell.CELL_TYPE_NUMERIC:
					                    if (DateUtil.isCellDateFormatted(copyCell)) {
					                    	cell.setCellValue(copyCell.getDateCellValue());
					                    } else {
					                    	cell.setCellValue(copyCell.getNumericCellValue());
					                    }
					                    break;
					                case Cell.CELL_TYPE_BOOLEAN:
					                	cell.setCellValue(copyCell.getBooleanCellValue());
					                    break;
					                case Cell.CELL_TYPE_FORMULA:
					                	cell.setCellValue(copyCell.getCellFormula());
					                    break;
					                default :
					                	cell.setCellValue(copyCell.getStringCellValue().trim());
							    }
							}  
						  }
						}
					}//end for
					createWB(addWb,fileReal+"(add)");
					JOptionPane.showMessageDialog(null,fileReal+"(add)."+_index+"完成");
				}
	}
	
	private void outType04(){
		 if(wbSheet!=null&&addWb!=null&&names!=null&&list!=null){
			 	Sheet sheet = addWb.getSheetAt(sheetNum_target);
			 	int numRow = sheet.getLastRowNum()+1;
				ArrayList<Integer> integers = list.get(0);
					for(int j=addJpanel.ignore_Rowtar;j<=wbSheet.getLastRowNum();j++){
							Row row = null;
							Row copy= null;					
					  if(j!=cellNum_target){
						 if((cellNum_target+1)>addJpanel.ignore_Rowtar)	
							row =sheet.createRow(j<cellNum_target?(j+numRow-addJpanel.ignore_Rowtar):(j+numRow-1-addJpanel.ignore_Rowtar));
						 else 
							row =sheet.createRow(j+numRow-addJpanel.ignore_Rowtar);
							copy= wbSheet.getRow(j);
					}
						if(copy!=null){
							for(int k=0;k<copy.getLastCellNum();k++){	
								Cell cell=null;
							 if(k>=addJpanel.ignore_Celltar)				 
								cell =row.createCell(integers.get((k-addJpanel.ignore_Celltar)));
							 else 
								cell = row.createCell(k);
								Cell copyCell = copy.getCell(k);
							if(copyCell!=null){	
							    switch (copyCell.getCellType()) {
					                case Cell.CELL_TYPE_STRING:
					                	cell.setCellValue(copyCell.getRichStringCellValue().getString().trim());
					                    break;
					                case Cell.CELL_TYPE_NUMERIC:
					                    if (DateUtil.isCellDateFormatted(copyCell)) {
					                    	cell.setCellValue(copyCell.getDateCellValue());
					                    } else {
					                    	cell.setCellValue(copyCell.getNumericCellValue());
					                    }
					                    break;
					                case Cell.CELL_TYPE_BOOLEAN:
					                	cell.setCellValue(copyCell.getBooleanCellValue());
					                    break;
					                case Cell.CELL_TYPE_FORMULA:
					                	cell.setCellValue(copyCell.getCellFormula());
					                    break;	   
					                default :
					                	cell.setCellValue(copyCell.getStringCellValue().trim());				      
							    }
							}  
						}	
						  }
					}//end for
					createWB(addWb,fileReal+"(add)");
					JOptionPane.showMessageDialog(null,fileReal+"(add)."+_index+"完成");
				}
	}
	
	
	private void createFolder(){
		String path = savePath;
		if(run.OStype==1){
			path +="/ExcelTool文件";
		if(type==1)
			path +="/"+filename;
			prolevel = 1;
			File file = new File(path);
			if(!file.exists()||!file.isDirectory()){
				file.mkdir();
			}
		}else if(run.OStype==2){
			path +="\\ExcelTool文件";
			if(type==1)
				path +="\\"+filename;
			prolevel = 2;
			File file = new File(path);
			if(!file.exists()||!file.isDirectory()){
				file.mkdir();
			}
		}	
	}
	
	private void createWB(Workbook workbook,String name){

			try {
				FileOutputStream out = null;
				String inString = ".xlsx";
				if(_index.equals("xls")){
					inString  = ".xls";
				}else if(_index.equals("xlsx")||_index.equals("")){
					inString = ".xlsx";
				}
			if(prolevel==1){
				out = new FileOutputStream(savePath+"/ExcelTool文件/"+name+inString);
			}else if(prolevel==2){
				out = new FileOutputStream(savePath+"\\ExcelTool文件\\"+name+inString);
			}
Log.info("excel is write！");
				workbook.write(out);
				out.flush();
				out.close();
			} catch (IOException e) {
				Log.warm(e.getCause().getMessage());
			}
		
	}
	
//	public static void main(String[] args) {
//		Workbook wb=new XSSFWorkbook();
//		Sheet sheet = wb.createSheet("new sheet");
//		CreationHelper helper = wb.getCreationHelper();
//	
//		Row row = (sheet).createRow(0);
//		Cell cell = row.createCell(0);
//		cell.setCellValue("吴轲是个傻逼");
//		
//		Cell cell2 = row.createCell(1);
//		cell2.setCellValue(new Date());
//			
//		CellStyle sellStyle = wb.createCellStyle();
//		sellStyle.setDataFormat(helper.createDataFormat().getFormat("m/d/yy h:mm"));
//		cell2.setCellStyle(sellStyle);
//		
//		try {
//			FileOutputStream out = new FileOutputStream("workbook2.xlsx");
//			wb.write(out);
//			out.close();
//		} catch (IOException e) {
//			Log.warm(e.getCause().getMessage());
//		}
//	}

}
