import java.util.List;

import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;

public class Sheet2ListHandler implements SheetContentsHandler {
    
	
    //collection 객체
    private List<String[]> rows;
    //collection 에 추가될 객체 startRow에서 초기화함.
    private String[] row;
    //collection 내 객체를 String[] 로 잡았기 때문에 배열의 길이를 생성시 받도록 설계
    private int columnCnt;
    //cell 이벤트 처리 시 해당 cell의 데이터가 배열 어디에 저장되야 할지 가리키는 pointer
    private int currColNum = 0;
    
    //외부 collection 과 배열 size를 받기 위해 추가한 부분입니다.
    public Sheet2ListHandler(
            List<String[]> rows
            , int columnsCnt
            ){        
        this.rows = rows;
        this.columnCnt = columnsCnt;
    }

	@Override
	public void startRow(int rowNum) {
		// TODO Auto-generated method stub
		  this.row = new String[columnCnt];
	        currColNum = 0;      
	}

	@Override
	public void endRow(int rowNum) {
		// TODO Auto-generated method stub
		 
        //cell 이벤트에서 담아놓은 row String[]를 collection에 추가
        //데이터가 하나도 없는 row는 collection에 추가하지 않도록 조건 추가
        boolean addFlag = false;
        for(String data : row){
            if(!"".equals(data))
                addFlag = true;
        }
        
        if(addFlag)rows.add(row);  
       
	}

	@Override
	public void cell(String cellReference, String formattedValue, XSSFComment comment) {
		// TODO Auto-generated method stub
	    //cell 이벤트 발생 시 해당 cell의 주소와 값을 받아옴. 
        //입맛에 맞게 처리하면됨.
		//cellReference 셀 이름 ex A1 B1
		
       row[currColNum++] = formattedValue == null ? "":formattedValue;
		
	}

	@Override
	public void headerFooter(String text, boolean isHeader, String tagName) {
		// TODO Auto-generated method stub
		
	}
}