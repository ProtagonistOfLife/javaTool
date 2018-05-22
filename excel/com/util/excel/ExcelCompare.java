package com.util.excel;


import java.io.File;
import java.io.IOException;
import java.util.Comparator;
import java.util.Map;
import java.util.TreeMap;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class ExcelCompare {
	private String oripath;
	private String propath;
	private Workbook oribook,probook;
	/*flag值为0：正常比较成
	 * 		  1：正常比较失败
	 * 		  2：异常比较
	 * */
	private int flag = 0;
	private int num = 0;
	private static Map<String, String> map = new TreeMap<>(new Comparator<String>() {
		@Override
		public int compare(String o1, String o2) {
			int n1 = Integer.parseInt(o1.substring(5));
			int n2 = Integer.parseInt(o2.substring(5));
			return n1-n2;
		}
		
	});
	
//	sheet内部表头对比
	public int sheetCompare(Sheet orisheet,Sheet prosheet){
		Cell[] oricells = orisheet.getRow(0).clone();
		Cell[] procells = prosheet.getRow(0).clone();
		if(oricells.length != procells.length || oricells.length < 1){
			map.put("point"+num++, "名为"+orisheet.getName()+"sheet为空或者两sheet内容不同");
			flag = 1;
		}
		for(int i = 0;i < oricells.length;i++){
			for(int j = 0;j < procells.length;j++){
				if(procells[j] != null && oricells[i].getContents().trim().equals(procells[j].getContents().trim())){
					flag = 1;
					oricells[i] = null;
					procells[j] = null;
					break;
				}
			}
		}
		StringBuilder sb = new StringBuilder();
		for(int i = 0;i < procells.length;i++){
			if(procells[i] != null){
				sb.append(procells[i].getContents()).append(",");
			}
		}
		if(sb.length() > 0){
			map.put("point"+num++, propath+"中"+prosheet.getName()+"独有的字段名："+sb.toString());
			sb = null;
		}
		sb = new StringBuilder();
		for(Cell c : oricells){
			if(c != null){
				sb.append(c.getContents()).append(",");
			}
		}
		if(sb.length() > 0){
			map.put("point"+num++, oripath+"中"+orisheet.getName()+"独有的字段名："+sb.toString());
			sb = null;
		}
		return flag;
		}
		
//	sheet名称对比
	public void getArrayDeffer(String[] oriarr,String[] proarr){
		String[] arr1 = oriarr.clone();
		String[] arr2 = proarr.clone();
		for(int i = 0;i < arr1.length;i++){
			for(int j = 0;j < arr2.length;j++){
				if(arr1[i].equals(arr2[j])){
					arr1[i] = null;
					arr2[j] = null;
					break;
				}
			}
		}
		StringBuilder sb = new StringBuilder();
		for(int i = 0;i < arr1.length;i++){
			if(arr1[i] != null){
				sb.append(arr1[i]).append(",");
			}
		}
		if(sb.length() > 0){
			map.put("point"+num++, oripath+"中独有的sheet名："+sb.toString());
			sb = null;
		}
		sb = new StringBuilder();
		for(String c : arr2){
			if(c != null){
				sb.append(c).append(",");
				System.out.println(c);
			}
		}
		if(sb.length() > 0){
			map.put("point"+num++, propath+"中独有的sheet名："+sb.toString());
			sb = null;
		}
		
	}
	
//	文件对比
	@SuppressWarnings("finally")
	public int fileCompare(String oripath,String propath){
		this.oripath = oripath;
		this.propath = propath;
		map.clear();
		num = 0;
		try {
			oribook = Workbook.getWorkbook(new File(oripath));
			probook = Workbook.getWorkbook(new File(propath));
			if(oribook.getNumberOfSheets() != probook.getNumberOfSheets()){
//				比较sheet数是否相同，不同为false
				flag = 1;
				map.put("point"+num, "sheet数量不一致");
				num++;
			}
			String[] orinames = oribook.getSheetNames();
			String[] pronames = probook.getSheetNames();
			getArrayDeffer(orinames, pronames);
		
			Sheet[] orisheets = oribook.getSheets();
			Sheet[] prosheets = probook.getSheets();
			for(int i = 0;i < orisheets.length;i++){
				for(int j = 0;j < prosheets.length;j++){
					if(orisheets[i].getName().equals(prosheets[j].getName()))
					sheetCompare(orisheets[i], prosheets[j]);
				}
			}
			oribook.close();
			probook.close();
		}catch (BiffException e) {
			flag = 2;
			e.printStackTrace();
		} catch (IOException e) {
			flag = 2;
			e.printStackTrace();
		}finally {
			if(oribook != null){
				oribook.close();
			}
			if(probook != null){
				probook.close();
			}
			return flag;
		}		
	}
	
	public Map<String, String> getMap(){
		return map;
	}
	
	/**
	 * xlsx读取
	 * @param args
	 */
	public boolean xlsxFileCompare(String oripath,String propath){
		return false;
	}
	
	public static void main(String[] args) {
		ExcelCompare ec = new ExcelCompare();
//		if(ec.fileCompare("D:/test1.xls", "D:/test2.xls") == 0){
		if(ec.fileCompare("D:/test3.xls","D:/SHIAUWCollect_2018-04-16.xls") == 0){
			System.out.println("相同");
		}else{
			System.out.println("不同");
		}
		System.out.println("---------------------------");
		if(map != null){
			for(Map.Entry<String, String> en : map.entrySet()){
				System.out.println(en.getKey() + "->" + en.getValue());
			}
		}
	}
}
