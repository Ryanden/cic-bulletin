import os
import requests
from bs4 import BeautifulSoup
from concurrent.futures import ThreadPoolExecutor

# 크롤링할 URL
url = "https://www.fustero.es/index_ja.php"

# 웹페이지 요청
session = requests.Session()  # 세션 사용으로 연결 재사용
response = session.get(url)
response.encoding = 'utf-8'

# 응답 확인
if response.status_code == 200:
    soup = BeautifulSoup(response.text, 'html.parser')

    # class가 'w3-container'와 'w3-row'를 모두 가진 section 요소를 선택합니다.
    sections = soup.find_all('section', class_=['w3-container', 'w3-row'])

    # 다운로드 폴더 생성
    download_dir = "download"
    os.makedirs(download_dir, exist_ok=True)

    # 다운로드 함수 정의
    def download_file(file_url):
        full_url = requests.compat.urljoin(url, file_url)
        file_name = os.path.join(download_dir, os.path.basename(file_url))

        print(f"Downloading: {full_url}")
        file_response = session.get(full_url, stream=True)

        if file_response.status_code == 200:
            with open(file_name, 'wb') as f:
                for chunk in file_response.iter_content(4096):  # 청크 크기 조정
                    f.write(chunk)
            print(f"Saved: {file_name}")
        else:
            print(f"Failed to download: {full_url}")

    # 다운로드할 파일 리스트 생성
    file_links = []
    for section in sections:
        links = section.find_all('a', href=True)
        for link in links:
            file_url = link['href']
            # if file_url.endswith(('.pptx', '.pdf', '.docx')):  # 필요한 파일 확장자 필터링
            if file_url.endswith(('.pptx')):  # 필요한 파일 확장자 필터링
                file_links.append(file_url)

    # 멀티스레딩 실행 (최대 5개 스레드 동시 실행)
    with ThreadPoolExecutor(max_workers=5) as executor:
        executor.map(download_file, file_links)

else:
    print("페이지 요청 실패:", response.status_code)
