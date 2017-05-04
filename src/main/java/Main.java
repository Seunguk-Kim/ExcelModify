import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;

public class Main {
	public static void main(String[] args) throws Exception {
		/** Parsing 설정 부분 **/

		// bulkInsertFile = File 객체 or FileInputStream (ex : new File(파일경로) 등으로
		// 넣을 수 있음)
		// OPCPackage 파일을 읽거나 쓸수있는 상태의 컨테이너를 생성함
		OPCPackage opc = OPCPackage.open("C:/Users/kim/Desktop/test.xlsx");
		OPCPackage opc2 = OPCPackage.open("C:/Users/kim/Desktop/code.xlsx");
		// opc 컨테이너 XSSF형식으로 읽어옴. 이 Reader는 적은 메모리로 sax parsing 을 하기 쉽게 만들어줌.
		XSSFReader xssfReader = new XSSFReader(opc);
		XSSFReader xssfReader2 = new XSSFReader(opc2);
		// XSSFReader 에서 sheet 별 collection으로 분할해서 가져옴.
		XSSFReader.SheetIterator itr = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
		XSSFReader.SheetIterator itr2 = (XSSFReader.SheetIterator) xssfReader2.getSheetsData();
		// 통합문서 내의 모든 Sheet에서 공유되는 스타일 테이블이라는데 정확한 사용용도는 모름;; ㅎㅎ
		StylesTable styles = xssfReader.getStylesTable();
		StylesTable styles2 = xssfReader.getStylesTable();
		// ReadOnlySharedStringsTable 이것도 정확한 역할은 모르겠음...ㅠ.ㅠ ...뭔가 data의 type을
		// 처리할 때 참조하는 것 같긴한데....
		ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(opc);
		ReadOnlySharedStringsTable strings2 = new ReadOnlySharedStringsTable(opc2);
		// 이건 내가 데이터를 파싱해서 담아올 List 객체...
		// 이것 마저도 메모리 부담이 된다면... Handler에 처리로직을 넣어서 한건씩 처리하면 됨.
		List<String[]> dataList = new ArrayList<String[]>();
		List<String[]> dataList2 = new ArrayList<String[]>();
		
		// Sheet 수 만큼 Loop를 돌린다.
		//while (itr.hasNext()) {
			InputStream sheetStream = itr.next();
			InputStream sheetStream2 = itr2.next();
			InputSource sheetSource = new InputSource(sheetStream);
			InputSource sheetSource2 = new InputSource(sheetStream2);
			// Sheet2ListHandler은 엑셀 data를 가져와서 SheetContentsHandler(Interface)를
			// 재정의 해서 만든 Class
			Sheet2ListHandler sheet2ListHandler = new Sheet2ListHandler(dataList,3);
			Sheet2ListHandler sheet2ListHandler2 = new Sheet2ListHandler(dataList2,3);
			// new XSSFSheetXMLHandler(StylesTable styles,
			// ReadOnlySharedStringsTable strings, SheetContentsHandler
			// sheet2ListHandler, boolean formulasNotResults)
			// formulasNotResults 이것도 무슨 옵션인지 모르겠음 ㅋㅋㅋ 이건 정보공유를 하겠다는 건지 ㅋㅋㅋ
			// 어쨌든 이 핸들러는 Sheet의 행(row) 및 Cell 이벤트를 생성합니다.
			ContentHandler handler = new XSSFSheetXMLHandler(styles, strings, sheet2ListHandler, false);
			ContentHandler handler2 = new XSSFSheetXMLHandler(styles2, strings2, sheet2ListHandler2, false);
			// sax parser를 생성하고...
			SAXParserFactory saxFactory = SAXParserFactory.newInstance();
			SAXParserFactory saxFactory2 = SAXParserFactory.newInstance();
			SAXParser saxParser = saxFactory.newSAXParser();
			SAXParser saxParser2 = saxFactory2.newSAXParser();
			// sax parser 방식의 xmlReader를 생성
			XMLReader sheetParser = saxParser.getXMLReader();
			XMLReader sheetParser2 = saxParser2.getXMLReader();
			// xml reader에 row와 cell 이벤트를 생성하는 핸들러를 설정한 후.
			sheetParser.setContentHandler(handler);
			sheetParser2.setContentHandler(handler2);
			// 위에서 Sheet 별로 생성한 inputSource를 parsing합니다.
			// 이 과정에서 handler는 row와 cell 이벤트를 생성하고 생성된 이벤트는 sheet2ListHandler 가
			// 받아서 처리합니다.
			// sheet2ListHandler의 내용은 아래를 참조하세요.
			sheetParser.parse(sheetSource);
			sheetParser2.parse(sheetSource2);
			
			SXSSFWorkbook workbook = new SXSSFWorkbook();
			
			SXSSFSheet sheet = workbook.createSheet("data");
			SXSSFSheet sheet2 = workbook.createSheet("code");
			
			int rownum2 = 0;
			for (int i = 0; i < dataList2.size(); i++) {
				Row row = sheet2.createRow(rownum2++);
				int cellnum2 = 0;
				for (int j = 0; j < dataList2.get(i).length; j++) {
					Cell cell = row.createCell(cellnum2++);
					cell.setCellValue(dataList2.get(i)[j]);
				}
			}
			
			
			int rownum = 0;
			for (int i = 0; i < dataList.size(); i++) {
				Row row = sheet.createRow(rownum++);
				int cellnum = 0;
				for (int j = 0; j < dataList.get(i).length; j++) {
				
					Cell cell = row.createCell(cellnum++);
					cell.setCellValue(dataList.get(i)[j]);

				}
				if(rownum==1){  // 첫행에 이름 추가 
					Cell cell = row.createCell(cellnum);
					cell.setCellValue("코드");
				}
				else{
					Cell cell = row.createCell(cellnum);
					System.out.println(rownum);
					String formula = "VLOOKUP(C"+rownum+",code!$A$2:$B$3,2,false)";
					System.out.println(formula);
					cell.setCellFormula(formula);
				}
					
			}//B열 2~5				
			FileOutputStream out = new FileOutputStream(new File("result.xlsx"));
			workbook.write(out);
			out.close();

			sheetStream.close();

	//	}

		opc.close();

	}
}
