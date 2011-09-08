package parser;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;

import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;

import parser.vo.DataVO;

import jxl.Cell;
import jxl.CellType;
import jxl.Sheet;
import jxl.Workbook;

/*
Copyright (c) 2011 xoombie.com

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

The Software shall be used for Good, not Evil.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
*/


/**
 * @author vikrambhatla
 * GameDataService class is converting Excel sheets to the *.json equivalent files
 * Its implemented to make the job of IPhone/IPad or Facebook Game Designers Job easy 
 * 
 * Every Excel sheet Should have START text from where the actual sheet data is getting started
 * Example GameData Excel file is in /assets/config
 */

public class GameDataService 
{

	public ArrayList<Object> parse(String path)
	{
		int i,j;
		String GAMEDATA_PATH = path;
		ArrayList<Object> data	= new ArrayList<Object>();
		File inputWorkBook	= new File(GAMEDATA_PATH);
		
		Workbook w;
		try{
			w = Workbook.getWorkbook(inputWorkBook);
			Sheet[] sheets	= w.getSheets();
			System.out.println("Sheet Content:"+sheets.length);
			
			for(i=0;i<sheets.length;i++){
				Sheet s	= sheets[i];
				
				int startRow	= 0;
				int startCol	= 0;
				
				String mySheetName = s.getName();
				
				for(j=0;j<s.getRows();j++){
					Cell cell	= s.getCell(startCol,j);
					if(cell.getContents().compareTo("START")==0){
						startRow	= j;
						break;
					}
				}
				
				ArrayList<Object> sheet	= new ArrayList<Object>();
				
				for(int row=startRow+2;row<s.getRows();row++){
					HashMap<String, Object> rowData	= new HashMap<String, Object>();
					for(int col=startCol+1;col<s.getColumns();col++){
						Cell cell	= s.getCell(col,row);
						String columnName	= ((Cell)s.getCell(col,startRow+1)).getContents();
						
						if(cell.getType()==CellType.LABEL)
							rowData.put(columnName, cell.getContents());
						else if(cell.getType()==CellType.NUMBER)
							rowData.put(columnName, new Float((cell.getContents())));
					}
					
					sheet.add(rowData);
				}
				DataVO sheetData	= new DataVO();
				sheetData.type		= mySheetName;
				sheetData.data		= sheet;
				System.out.println("Sheet Content:"+sheet);
				
				data.add(sheetData);
			}
			w.close();
		}catch(Exception e){
			System.out.println("Gamedata Read Error:"+e.getMessage());
			e.printStackTrace();
		}
		System.out.println("Sheet Content:"+data);
		return data;
	}
	
	public void searilizeToJson(String path, String file)
	{
		String JSON_FILE;
		ArrayList<Object> gamedata	= parse(path+file);
		int i = 0, j=0;
		DataVO sheetData;
		ArrayList<Object> sheet;
		HashMap<String, Object> rowData;
		
		JSONArray sheetJsonArray;
		JSONObject sheetJsonObject = null;
		
		Iterator<String> iterator;
		
		for(i=0;i<gamedata.size();i++)
		{
			sheetData = (DataVO)gamedata.get(i);
			JSON_FILE = path+ (sheetData.type+".json");
			sheet = sheetData.data;
			
			sheetJsonArray = new JSONArray();
			
			for(j=0;j<sheet.size();j++)
			{
				rowData = (HashMap<String, Object>)sheet.get(j);
				
				 iterator = rowData.keySet().iterator();
			     sheetJsonObject = new JSONObject();
			    while (iterator.hasNext()) {
			      String key = (String) iterator.next();
			      
			      try
			      {
			    	  sheetJsonObject.put(key,rowData.get(key));
			      }catch(JSONException e){}
			    }				
			    
			    sheetJsonArray.put(sheetJsonObject);
			    
			    try
			    {
				    FileWriter fstream = new FileWriter(JSON_FILE);
			        BufferedWriter out = new BufferedWriter(fstream);
			        out.write(sheetJsonArray.toString());
			        out.close();
		        }
			    catch(IOException e){
					System.out.println(JSON_FILE+"File Read Error:"+e.getMessage());
			    	
			    }
			}
			
		}
		
	}
	
}
