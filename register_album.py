import os
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

LOGIN_URL = "https://www.mims.or.kr/login"
ALBUM_REGISTER_URL = "https://www.mims.or.kr/mypage/meta"


def login(driver, mims_id: str, mims_password: str) -> bool:
    driver.get(LOGIN_URL)
    try:
        wait = WebDriverWait(driver, 10)
        id_input = wait.until(EC.presence_of_element_located((By.ID, "inputEmail")))
        pw_input = wait.until(EC.presence_of_element_located((By.ID, "inputPwd")))
        login_button = wait.until(EC.element_to_be_clickable((By.ID, "login-btn")))

        id_input.send_keys(mims_id)
        pw_input.send_keys(mims_password)
        login_button.click()

        # 로그인 성공 신호: 상단 내비의 My앨범 링크 존재
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'a[href="/mypage/album"]'))
        )
        print("로그인 성공!")
        return True
    except Exception as e:
        print(f"로그인 실패: {e}")
        return False


def goto_album_register(driver) -> bool:
    try:
        # 더 안정적인 방식: 직접 URL 이동
        driver.get(ALBUM_REGISTER_URL)

        # URL 확인 및 폼/입력 존재 대기(있으면 좋고, 없으면 URL 기준으로 통과)
        WebDriverWait(driver, 10).until(EC.url_contains("/mypage/meta"))
        try:
            WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "form, input, select, textarea"))
            )
        except Exception:
            pass

        print("앨범등록 페이지 진입 완료!")

        # 대량등록(엑셀) 버튼 클릭 및 패널 오픈 대기
        bulk_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "register-excel-btn"))
        )
        bulk_btn.click()
        excel_card = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "excel-card"))
        )
        WebDriverWait(driver, 10).until(lambda d: 'd-none' not in excel_card.get_attribute('class'))
        print("대량등록(엑셀) 패널 열기 완료!")

        # 숨겨진 파일 입력을 찾아 노출 후, 지정 파일 선택(send_keys)
        file_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "mims-excel-upload"))
        )

        # 절대 경로 생성 (현재 작업 디렉토리 기준)
        excel_filename = "박성태 - 기도왕｜MIMS.xlsx"
        excel_path = os.path.abspath(os.path.join(os.getcwd(), excel_filename))
        if not os.path.exists(excel_path):
            raise FileNotFoundError(f"엑셀 파일을 찾을 수 없습니다: {excel_path}")

        # display:none 회피: 입력을 보이도록 만들고 값 전송
        driver.execute_script(
            "arguments[0].classList.remove('d-none'); arguments[0].style.display='block';",
            file_input,
        )
        file_input.send_keys(excel_path)
        print(f"파일 선택 완료: {excel_path}")

        # 큐에 행이 나타날 때까지 대기 (업로드 준비됨)
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "#excel-card tbody tr"))
        )

        # 업로드 버튼 클릭 (fa-upload 아이콘 포함 버튼)
        upload_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//div[@id='excel-card']//button[.//span[contains(@class,'fa-upload')]]"))
        )
        upload_btn.click()
        print("업로드 버튼 클릭 완료. 업로드 진행 대기...")

        # 업로드 성공 배지 등장 대기
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "#excel-card .badge-success"))
        )
        print("업로드 완료 감지!")

        return True
    except Exception as e:
        print(f"앨범등록 페이지/업로드 흐름 실패: {e}")
        return False


def main():
    load_dotenv()
    mims_id = os.getenv("MIMS_ID")
    mims_password = os.getenv("MIMS_PASSWORD")

    if not mims_id or not mims_password:
        print("환경변수 MIMS_ID/MIMS_PASSWORD가 설정되지 않았습니다. .env를 확인하세요.")
        return

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

    try:
        if not login(driver, mims_id, mims_password):
            return
        goto_album_register(driver)
        input("작업 완료. Enter를 누르면 브라우저를 닫습니다...")
    finally:
        driver.quit()


if __name__ == "__main__":
    main()
