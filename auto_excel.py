import openpyxl
from bs4 import BeautifulSoup
import requests
import time

# 엑셀 파일 열기
workbook = openpyxl.load_workbook('url.xlsx')
sheet = workbook.active

# 블로그 링크 리스트 (비교 대상)
target_links = [
    "@@@@@"
]

for row in sheet.iter_rows(min_row=2, max_col=1):  
    search_url = row[0].value  
    row_index = row[0].row  

    if search_url: 
        
        
        try:
            response = requests.get(search_url)
            response.raise_for_status()  
            soup = BeautifulSoup(response.text, 'html.parser')
        except requests.exceptions.RequestException as e:
            print(f"검색 중 오류 발생: {e}")
            sheet[f"B{row_index}"] = "검색 오류"  
            continue  

        # 광고 제외
        hrefs = []
        count = 0
        for a_tag in soup.find_all('a', class_='name'):
            if not a_tag.find_parent(class_='link_ad'):
                href = a_tag['href']
                # 광고 링크 제외
                if not href.startswith('https://adcr.naver.com/'):
                    hrefs.append(href)
                    count += 1
                    if count == 40:  # 상위n개의 블로그를 확인
                        break

        # 타겟 링크가 몇 번째에 있는지 확인
        rankings = {}
        for target_link in target_links:
            if target_link in hrefs:
                rankings[target_link] = hrefs.index(target_link) + 1  # 순위 저장

        # 가장 높은 등수의 링크 찾기
        if rankings:
            best_link = min(rankings, key=rankings.get)  # 가장 작은 값을 가진 값
            best_rank = rankings[best_link]
            result = best_rank
        else:
            result = "타겟 링크 없음"

        
        sheet[f"B{row_index}"] = result
        print(f"B{row_index}에 결과 저장: {result}")
        time.sleep(2)

workbook.save('urls_with_results.xlsx')
print("엑셀 파일에 검색 결과가 저장되었습니다.")
