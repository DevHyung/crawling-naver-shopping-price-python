from bs4 import BeautifulSoup
from openpyxl import Workbook
import requests

def saveExcel(query,data):
    """
    :param query: naver 쇼핑물에 query 한걸로 파일명 을 만들꺼
    :param data:  크롤링한 결과물
    :return:  NONE
    """
    # 엑셀시트 header 설정 및, 열의 넓이 설정
    header1 = ['품명','최저가','링크']
    wb = Workbook()
    ws1 = wb.worksheets[0]
    ws1.column_dimensions['A'].width = 50
    ws1.column_dimensions['B'].width = 20
    ws1.column_dimensions['C'].width = 50
    # 데이터 삽입
    # itemlist 가 [품명,최저가,링크] 이런식으로 온걸
    # openpyxl 객체 ws1 에 append 시키면 들어감
    for itemlist in data:
        ws1.append(itemlist)
    wb.save(query+".xlsx")

if __name__ == "__main__": # 직접실행시키는 경우
    #검색할 단어를 입력받고
    query = input("검색할 단어를 입력하세요 : ")
    # [ [품명,최저가,링크], [품명,최저가,링크] ]  이런 구조를 지닌 2차원 배열을 만들변수
    datalist = []
    # 페이지에 GET 요청을 한후 소스코드를 받아옴
    html = requests.get('https://search.shopping.naver.com/search/all.nhn?where=all&frm=NVSCTAB&query='+query)
    # 그코드를 python 에서 분석하기 쉽게 BeautifulSoup 객체로 변환
    bs4 = BeautifulSoup(html.text,'lxml')
    # https://search.shopping.naver.com/search/all.nhn?query=gtx1070&cat_id=&frm=NVSHATC 기준
    # item들은 <li class="_mod..." 이렇게 감싸져있는걸 모두 가져옴
    itemlist = bs4.find_all("li",class_="_model_list _itemSection")
    for item in itemlist:
        # 각 태그에 해당하는 정보를 꺼내옴
        itemdiv = item.find('div', class_="info")
        title = itemdiv.find('a').get_text().strip()
        link = itemdiv.find('a')['href']
        price = itemdiv.find('span',class_='price').find('span',class_='num _price_reload').get_text().strip()
        #datalist 배열에 title,price,link 를 배열로 감싸서 계속 추가
        # 결과적으로 datalist는 1차원배열들을 집합으로 가지고있는 2차원 배열이됌
        datalist.append( [title,price,link] )
    #저장
    saveExcel(query,datalist)
