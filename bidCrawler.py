# 크롬 브라우저를 띄우기 위해, 웹드라이버를 가져오기
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from openpyxl import Workbook

# 크롬 드라이버로 크롬을 실행한다.
driver = webdriver.Chrome('./chromedriver')

try:
    # query list
    query_dict = ['발사체', '로켓', '추진', '위성', '우주', 'UAV', 'satellite', '무인기', '방사선', '방사능', '원자력', '선량', '핵종', '감마', '지상', '영상', '기상', 'GRPS', '425', '딥러닝', '인공지능', '머신러닝', '기계학습', 'AI', 'SBAS', 'KASS', 'KCS', '통합운영국', '다종', '다중', '수신시스템', '처리시스템', 'KPS', '대구경', '다중대역']

    # for each query_dict iterate

    #Workbook 생성
    wb = Workbook()
    ws = wb.active

    for i in query_dict:
        # 입찰정보 검색 페이지로 이동
        driver.get('https://www.g2b.go.kr:8101/ep/tbid/tbidFwd.do')
        # 업무 종류 체크 
        
        # 검색어
        query = i
        # id값이 bidNm인 태그 가져오기
        bidNm = driver.find_element(By.ID, 'bidNm')
        # 내용을 삭제 (버릇처럼 사용할 것!)
        bidNm.clear()
        # 검색어 입력후 엔터
        bidNm.send_keys(query)
        bidNm.send_keys(Keys.RETURN)

        # 기간
        strdate = '2022/12/14'
        enddate = '2022/12/15'
        # id값이 fromBidDt인 태그 가져오기
        fromBidDt = driver.find_element(By.ID, 'fromBidDt')
        # 내용 삭제
        fromBidDt.clear()
        # 기간 입력후 엔터
        fromBidDt.send_keys(strdate)
        fromBidDt.send_keys(Keys.RETURN)

        # id값이 toBidDt인 태그 가져오기
        toBidDt = driver.find_element(By.ID, 'toBidDt')
        # 내용 삭제
        toBidDt.clear()
        # 기간 입력후 엔터
        toBidDt.send_keys(enddate)
        toBidDt.send_keys(Keys.RETURN)

        # 검색 조건 체크 ('검색기간 1달': 'setMonth1_1')
        option_dict = {'입찰마감건 제외': 'exceptEnd', '검색건수 표시': 'useTotalCount'}
        for option in option_dict.values():
            checkbox = driver.find_element(By.ID, option)
            checkbox.click()

        # 목록수 100건 선택 (드롭다운)
        recordcountperpage = driver.find_element(By.NAME, 'recordCountPerPage')
        selector = Select(recordcountperpage)
        selector.select_by_value('100')

        # 검색 버튼 클릭
        search_button = driver.find_element(By.CLASS_NAME, 'btn_mdl')
        search_button.click()

        # 검색 결과 확인
        elem = driver.find_element(By.CLASS_NAME, 'results')
        div_list = elem.find_elements(By.TAG_NAME, 'div')

        # 검색 결과 모두 긁어서 리스트로 저장
        results = []
        for div in div_list:
            results.append(div.text)
            a_tags = div.find_elements(By.TAG_NAME, 'a')
            if a_tags:
                for a_tag in a_tags:
                    link = a_tag.get_attribute('href')
                    results.append(link)

        # 검색결과 모음 리스트를 12개씩 분할하여 새로운 리스트로 저장 
        result = [results[i * 12:(i + 1) * 12] for i in range((len(results) + 12 - 1) // 12 )]

        print (result)

        '''
        for i in results:
            ws.append([i])
        wb.save("g2b_results.xlsx")
        '''

        # 리스트에서 특정 항목만 가져와서 엑셀 파일에 저장
        for a,b,c,d,e,f,g,h,i,j,k,l in result:
            print([c,e,g])
            ws.append([c])
            ws.append([e])
            ws.append([g])
            ws.append([])
        wb.save("g2b_results.xlsx")

except Exception as e:
    # 위 코드에서 에러가 발생한 경우 출력
    print(e)
finally:
    # 에러와 관계없이 실행되고, 크롬 드라이버를 종료
    driver.quit()