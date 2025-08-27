import os
import time
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, UnexpectedAlertPresentException
from selenium.webdriver.common.keys import Keys
import traceback

def login(driver, mims_id, mims_password):
    driver.get("https://www.mims.or.kr/login")
    try:
        wait = WebDriverWait(driver, 10)
        id_input = wait.until(EC.presence_of_element_located((By.ID, "inputEmail")))
        pw_input = wait.until(EC.presence_of_element_located((By.ID, "inputPwd")))
        login_button = wait.until(EC.element_to_be_clickable((By.ID, 'login-btn')))

        id_input.send_keys(mims_id)
        pw_input.send_keys(mims_password)
        login_button.click()

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'a[href="/mypage/album"]'))
        )
        print("로그인 성공!")
        return True
    except Exception as e:
        print(f"로그인 실패: {e}")
        return False

def find_approved_albums(driver):
    try:
        # 'My앨범' 페이지로 바로 이동
        my_album_link = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[href="/mypage/album"]'))
        )
        my_album_link.click()

        # 앨범 목록 컨테이너가 로드될 때까지 대기
        album_container_selector = "div.mims-pmb"
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, album_container_selector))
        )

        approved_albums = []
        album_cards = driver.find_elements(By.CSS_SELECTOR, f"{album_container_selector} .thumbnail-style")
        print(f"현재 페이지에서 {len(album_cards)}개의 앨범을 찾았습니다.")

        for card in album_cards:
            try:
                title_element = card.find_element(By.CSS_SELECTOR, "h3 > a")
                album_title = title_element.text
                
                album_link_element = card.find_element(By.CSS_SELECTOR, "a.go-view")
                album_code = album_link_element.get_attribute('data-album-code')

                approved_albums.append({"title": album_title, "code": album_code, "element": album_link_element})
                print(f"앨범 찾음: {album_title} (코드: {album_code})")
            except Exception as e:
                print(f"앨범 정보를 가져오는 중 오류 발생: {e}")
        
        return approved_albums

    except Exception as e:
        print(f"앨범을 찾는 중 오류 발생: {e}")
        return []

def issue_codes(driver):
    print("앨범 상세 페이지로 이동했습니다. ISRC/UCI 코드 확인 및 발급을 시작합니다.")

    def extract_codes(driver):
        codes_list = []
        try:
            # 'ISRC/Music.UCI'라는 고유한 헤더를 기준으로 수록곡 목록을 찾는 가장 확실한 방식으로 경로를 수정했습니다.
            track_list_xpath = "//th[contains(text(), 'ISRC/Music.UCI')]/ancestor::table/tbody/tr"
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, track_list_xpath)))
            track_rows = driver.find_elements(By.XPATH, track_list_xpath)

            if not track_rows: return []

            for row in track_rows:
                isrc_elements = row.find_elements(By.CSS_SELECTOR, "span.g-bg-darkred")
                if not isrc_elements:
                    return None  # ISRC not found, indicates codes need issuing

                title = row.find_element(By.CSS_SELECTOR, "td:nth-child(4) a").text
                isrc = isrc_elements[0].get_attribute("data-clipboard-data")
                uci = row.find_element(By.CSS_SELECTOR, "span.g-bg-blue").get_attribute("data-clipboard-data")
                codes_list.append({"title": title, "isrc": isrc, "uci": uci})
            return codes_list
        except Exception:
            return None

    def _count_rows(driver):
        try:
            track_list_xpath = "//th[contains(text(), 'ISRC/Music.UCI')]/ancestor::table/tbody/tr"
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, track_list_xpath)))
            return len(driver.find_elements(By.XPATH, track_list_xpath))
        except Exception:
            return 0

    def _wait_for_isrc_applied(driver, expected_rows):
        # 우선 하나 이상 나타날 때까지 대기
        WebDriverWait(driver, 30).until(
            lambda d: len(d.find_elements(By.CSS_SELECTOR, "span.g-bg-darkred[data-clipboard-data]")) >= 1
        )
        # 가능하면 전체 행 수만큼 채워질 때까지 추가 대기 (최대 20초)
        try:
            WebDriverWait(driver, 20).until(
                lambda d: len(d.find_elements(By.CSS_SELECTOR, "span.g-bg-darkred[data-clipboard-data]")) >= expected_rows
            )
        except TimeoutException:
            pass

    def _wait_for_uci_applied(driver, expected_rows):
        WebDriverWait(driver, 30).until(
            lambda d: len(d.find_elements(By.CSS_SELECTOR, "span.g-bg-blue[data-clipboard-data]")) >= 1
        )
        try:
            WebDriverWait(driver, 20).until(
                lambda d: len(d.find_elements(By.CSS_SELECTOR, "span.g-bg-blue[data-clipboard-data]")) >= expected_rows
            )
        except TimeoutException:
            pass

    def _count_isrc_applied(driver):
        return len(driver.find_elements(By.CSS_SELECTOR, "span.g-bg-darkred[data-clipboard-data]"))

    def _count_uci_applied(driver):
        return len(driver.find_elements(By.CSS_SELECTOR, "span.g-bg-blue[data-clipboard-data]"))

    def _accept_all_alerts(driver, label: str, max_tries: int = 5):
        for i in range(1, max_tries + 1):
            try:
                WebDriverWait(driver, 2).until(EC.alert_is_present())
                alert = driver.switch_to.alert
                try:
                    print(f"[{label} ALERT #{i}] {alert.text}")
                except Exception:
                    pass
                alert.accept()
                time.sleep(0.2)
            except TimeoutException:
                break
            except Exception as e:
                print(f"[{label} ALERT HANDLER ERROR] {e}")
                print(traceback.format_exc())
                break

    def _click_and_accept(driver, by, ident, label):
        try:
            btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((by, ident)))
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
            driver.execute_script("arguments[0].click();", btn)
            # 즉시 뜨는 알럿(확인/결과) 모두 수락
            _accept_all_alerts(driver, label, max_tries=5)
            print(f"{label} 발급 확인 완료. 페이지 반영 대기...")
            # 대기 중 발생하는 지연 알럿까지 처리하며 재시도
            for _ in range(3):
                try:
                    WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, "//th[contains(text(), 'ISRC/Music.UCI')]"))
                    )
                    break
                except UnexpectedAlertPresentException:
                    _accept_all_alerts(driver, label, max_tries=5)
                except TimeoutException:
                    # 혹시 늦게 뜬 알럿 있을 수 있어 한 번 더 수락 시도
                    _accept_all_alerts(driver, label, max_tries=1)
            return True
        except NoSuchElementException as e:
            print(f"{label} 발급 버튼을 찾을 수 없습니다. {e}")
            print(traceback.format_exc())
            return False
        except TimeoutException as e:
            print(f"{label} 발급 확인창이 나타나지 않았거나 페이지 로딩에 실패했습니다. {e}")
            print(traceback.format_exc())
            return False

    try:
        # 수록곡 섹션 로드 및 총 행수 파악
        total_rows = _count_rows(driver)
        if total_rows == 0:
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, "//th[contains(text(), 'ISRC/Music.UCI')]/ancestor::table/tbody/tr"))
            )
            total_rows = _count_rows(driver)

        # 1) 현재 적용 상태 파악
        isrc_applied = _count_isrc_applied(driver)
        uci_applied = _count_uci_applied(driver)
        need_isrc = isrc_applied < total_rows
        need_uci = uci_applied < total_rows

        if not need_isrc and not need_uci:
            print("ISRC/UCI가 모두 존재합니다. 코드만 추출합니다.")
            return extract_codes(driver)

        # 2) 필요한 것만 발급
        did_issue = False
        if need_isrc:
            print("ISRC 발급 버튼 클릭...")
            if _click_and_accept(driver, By.ID, "setTrackIsrc", "ISRC"):
                _wait_for_isrc_applied(driver, total_rows)
            did_issue = True

        if need_uci:
            print("UCI 발급 버튼 클릭...")
            if _click_and_accept(driver, By.ID, "setTrackUCI", "UCI"):
                _wait_for_uci_applied(driver, total_rows)
            did_issue = True

        # 3) 반영 대기 후 1초 지연 → 새로고침 1회 → 추출 → 마지막에 새로고침 1회 더 → 최종 추출
        if did_issue:
            time.sleep(1)
            driver.refresh()
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, "//th[contains(text(), 'ISRC/Music.UCI')]/ancestor::table/tbody/tr"))
            )
            first_codes = extract_codes(driver)

            driver.refresh()
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, "//th[contains(text(), 'ISRC/Music.UCI')]/ancestor::table/tbody/tr"))
            )
            print("코드 발급/이중 새로고침 후 최종 코드 추출을 시도합니다.")
            final_codes = extract_codes(driver)
            return final_codes or first_codes

        # 발급이 없었으면 즉시 추출
        return extract_codes(driver)

    except Exception as e:
        print(f"코드 확인/발급 중 예상치 못한 오류 발생: {e}")
        print(traceback.format_exc())
        return None

def main():
    load_dotenv()

    MIMS_ID = os.getenv("MIMS_ID")
    MIMS_PASSWORD = os.getenv("MIMS_PASSWORD")

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    
    if login(driver, MIMS_ID, MIMS_PASSWORD):
        albums = find_approved_albums(driver)
        if albums:
            # 가장 첫 번째 (최신) 앨범의 상세 페이지로 이동
            latest_album_code = albums[0]["code"]
            latest_album_title = albums[0]["title"]
            album_url = f"https://www.mims.or.kr/mypage/view/album/{latest_album_code}"
            
            print(f"\n가장 최신 앨범 '{latest_album_title}'의 상세 페이지로 이동합니다.")
            print(f"URL: {album_url}")
            driver.get(album_url)
            
            # 페이지 전환 후 '수록곡' 섹션의 실제 내용(테이블의 첫 번째 줄)이 나타날 때까지 명시적으로 기다립니다.
            try:
                # 'ISRC/Music.UCI'라는 고유한 헤더를 기준으로 수록곡 목록을 찾는 가장 확실한 방식으로 경로를 수정했습니다.
                track_list_xpath = "//th[contains(text(), 'ISRC/Music.UCI')]/ancestor::table/tbody/tr"
                WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.XPATH, track_list_xpath))
                )
                issued_codes = issue_codes(driver)
    
                if issued_codes:
                    print("\n--- 코드 추출 결과 ---")
                    for item in issued_codes:
                        print(f"  곡명: {item['title']}")
                        print(f"  ISRC: {item['isrc']}")
                        print(f"  UCI:  {item['uci']}")
                        print("-" * 22)
                else:
                    print("앨범에서 ISRC/UCI 코드를 찾지 못했습니다. (발급 버튼이 없었거나, 발급에 실패했을 수 있습니다.)")
            except Exception as e:
                print(f"코드 추출 과정에서 오류가 발생했습니다: {e}")

        else:
            print("페이지에서 앨범을 찾지 못했습니다.")

    input("Press Enter to quit...")
    driver.quit()

if __name__ == "__main__":
    main()
