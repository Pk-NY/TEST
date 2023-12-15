[일간 공연랭킹알아보기]

예스24 티켓 사이트에서 일간 공연 랭킹을 추출하여 결과 파일 생성하는 것을 REFramework로 구현

- 예스24티켓 공연 카테고리 리스트를 가져와서 카테고리별로 1~50위 랭킹 데이터 추출 엑셀에 카테고리별 시트를 추가하여 결과 파일 생성

[사전작업]

Input 폴더 : 엑셀템플릿 작성, VBA 코드 파일

Config.xlsx : ResultFileName(결과파일이름), SiteURL(사이트주소), TemplateFile(적용할 엑셀양식경로),Category(가져올 공연카테고리), vba코드파일 경로
             등 작성

 [개인작업파일]
 
 Business 폴더: 
 
 CheckOutputFolder : output폴더 확인 후 생성
 
 CopyTemplate: 템플릿파일 복사 후 결과파일 생성
 
 ExcelMacro: 결과파일에 vba 코드 적용

[작업과정]

1. **INITIALIZE PROCESS**

   output폴더확인 후 생성, 템플릿파일복사결과파일 생성, yes24공연랭킹 사이트 접속해서 공연카테고리, URL DT로 생성
   config에서 가져올 공연 카테고리 읽어와서 해당 카테고리 데이터를 dt_TransactionData에 담기

2. **GET TRANSACTION DATA**
 
   dt_TransactionData : 공연카테고리,URL
   TransactionItem 타입 : DataRow

3. **PROCESS TRANSACTION**

    yes24티켓 사이트에서 순위, 공연명, 기간과 장소를 추출 후 카테고리별로 엑셀에 작성
   - 1위~3위 데이터와 4위부터의 DT를 따로 추출 후 합해 준다.
     
4. **END PROCESS**

   결과파일에 vba 코드 적용하여 보기 편하게 정리, 사이트가 열려 있는지 확인 후 닫기


