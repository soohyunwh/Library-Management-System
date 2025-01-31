package Book;

//*import org.apache.poi.ss.usermodel
//Apache POI 라이브러리에서 제공하는 엑셀 작업 관련 클래스들
//*org.apache.poi.xssf.usermodel
// .xlsx 파일 형식을 다루는 클래스들

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.Scanner;

public class book {
    public static void main(String[] args) {
        //1.엑셀파일 경로설정
        String file = "D:\\Workdir\\book.xlsx"; //엑셀파일 경로

        //2.도서 데이터를 저장할 map 생성(책 제목 -> "장소:저자")
        //여기서 키는 책 제목이고, 값은 "장소: 저자" 형태의 문자열
        Map<String, String> bookMap = new HashMap<>();

        //3.엑셀파일 읽기 시도 (try-with-resources 사용)
        // try-with-resources는 파일이나 리소스를 자동으로 닫아주는 구조
        try (FileInputStream fis = new FileInputStream(file); // FileInputStream으로 파일 읽기
             Workbook workbook = new XSSFWorkbook(fis)) { // .xlsx 형식의 엑셀 파일을 읽기 위해 Workbook 객체 생성


            // workbook의 모든 시트를 반복
            for (Sheet sheet : workbook) {
                String location = sheet.getSheetName(); // 시트 이름을 읽음(이름은 책이 위치한 장소)

                //시트 내의 모든 행(Row)를 반복
                for (Row row : sheet) {
                    if (row.getRowNum() == 0) continue; // 첫번째 행은 데이터가 아니므로 건너뛰기

                    //각 행에서 셀을(cell)을 읽음
                    Cell bookNameCell = row.getCell(1); // 첫번째 열(셀): 책 제목
                    Cell authorCell = row.getCell(2); // 두번째 열(셀): 저자

                    //빈 셀일 경우 건너뛰기
                    if (bookNameCell == null || authorCell == null) continue;

                    //셀 값을 문자열로 변환
                    //getStringCellValue(): 셀에 있는 값을 문자열로 반환
                    String bookName = getCellValueAsString(bookNameCell).trim(); // 책 제목 (앞 뒤 공백 제거)
                    String author = getCellValueAsString(authorCell).trim();

                    //Map에 데이터 저장
                    // 책 제목을 키, 장소: 저자를 값으로 저장
                    bookMap.put(bookName, location + ": " + author);

                }
            }
        } catch (IOException e) { // 파일 입출력 중 발생할 수 있는 예외처리
            System.out.println("엑셀파일을 읽는 중 오류가 발생했습니다"); // 오류메세지 입력
            e.printStackTrace(); // 오류 상세 정보를 출력(디버깅용)
            return; //프로그램 종료
        }
        //4. 사용자 입력받기
        Scanner scanner = new Scanner(System.in);
        System.out.println("검색할 책 제목을 입력하세요");
        String searchBook = scanner.nextLine().trim(); // 사용자가 입력한 책 제목을 읽고 공백 제거

        //5. 책 검색
        if (bookMap.containsKey(searchBook)) { //사용자가 입력한 책 제목이 Map의 키에 있는지 확인

            String bookInfo = bookMap.get(searchBook);

            String[] infoparts = bookInfo.split(": ", 2);

            String location = infoparts[0];

            String authors = infoparts.length > 1 ? infoparts[1].replace(";","/") :"저자정보 없음";
            //검색결과가 있을경우 출력
            System.out.println("\n" + searchBook + "의 위치와 저자 정보: ");
            System.out.println(location);
            System.out.println("저자: "+authors);
            //Map의 get()메서드로 해당 책 제목의 장소와 저자 정보를 가져옴
        } else {
            //검색결과가 없을경우
            System.out.println(searchBook + "은(는) 목록에 없습니다");
        }
    }
    private static String getCellValueAsString(Cell cell) {
        if (cell == null) return  "";

        switch (cell.getCellType()) {
            case STRING:
                return  cell.getStringCellValue().trim();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                }
                return String.valueOf((int)cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }
}
