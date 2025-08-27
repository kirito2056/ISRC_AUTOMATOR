import os
import time
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException
from openpyxl import load_workbook
import datetime as dt
from selenium.webdriver.common.keys import Keys

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

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'a[href="/mypage/album"]'))
        )
        print("로그인 성공!")
        return True
    except Exception as e:
        print(f"로그인 실패: {e}")
        return False


def _get_field_value(driver, field_id: str) -> str:
    try:
        el = driver.find_element(By.ID, field_id)
        tag = el.tag_name.lower()
        if tag in ("input", "textarea"):
            return (el.get_attribute("value") or "").strip()
        if tag == "select":
            try:
                return el.find_element(By.CSS_SELECTOR, "option:checked").text.strip()
            except Exception:
                return ""
        return (el.text or "").strip()
    except Exception:
        return ""


def _check_required_and_go_next(driver) -> None:
    WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.ID, "meta-album-form"))
    )
    form = driver.find_element(By.ID, "meta-album-form")
    labels = form.find_elements(By.CSS_SELECTOR, "label.required")

    missing = []
    for lb in labels:
        field_id = lb.get_attribute("for") or ""
        if not field_id:
            continue
        val = _get_field_value(driver, field_id)
        if field_id == "albumTitle" and not val:
            try:
                val = driver.find_element(By.ID, "display_album_title").text.strip()
            except Exception:
                pass
        if not val:
            label_text = (lb.text or field_id).strip()
            missing.append((field_id, label_text))

    if missing:
        print("필수 입력 누락으로 앨범정보에서 이동 중단:")
        for fid, ltxt in missing:
            print(f" - {ltxt} (id={fid}) 비어있음")
        return

    try:
        next_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "meta-next-album-btn"))
        )
        next_btn.click()
        print("곡정보 탭으로 이동 중...")
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, "track-list"))
        )
        print("곡정보 탭 로딩 완료")
    except Exception as e:
        print(f"곡정보 이동 실패: {e}")


def _check_track_required_and_next(driver) -> None:
    # 곡정보 탭 폼/테이블 대기
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "meta-tracks"))
        )
    except TimeoutException:
        # 이미 다른 탭일 수 있으므로 조용히 리턴
        return

    required_ids = ["diskNo", "trackNo", "trackTitle", "trackArtistName", "trackGenre", "duration"]
    missing = []

    # duration은 hidden #duration 또는 HH/MM/SS 조합 확인
    for fid in required_ids:
        if fid == "duration":
            dur_hidden = _get_field_value(driver, "duration")
            if dur_hidden:
                continue
            hh = _get_field_value(driver, "duration_hh")
            mm = _get_field_value(driver, "duration_mm")
            ss = _get_field_value(driver, "duration_ss")
            if not (hh and mm and ss):
                missing.append((fid, "재생시간"))
            continue
        val = _get_field_value(driver, fid)
        if not val:
            label_text = fid
            try:
                label_text = driver.find_element(By.CSS_SELECTOR, f"label[for='{fid}']").text.strip()
            except Exception:
                pass
            missing.append((fid, label_text))

    if missing:
        print("필수 입력 누락으로 곡정보에서 이동 중단:")
        for fid, ltxt in missing:
            print(f" - {ltxt} (id={fid}) 비어있음")
        return

    # 저장 후 다음으로 버튼 클릭
    try:
        save_next_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "update-meta-track-next-track-btn"))
        )
        save_next_btn.click()
        print("곡정보 저장 후 다음으로 이동 중...")
        # 다음 탭(앨범중복확인)에서 버튼/컨텐츠 대기
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "meta-next-right-new-btn"))
        )
        print("앨범중복확인 탭 로딩 완료")
    except Exception as e:
        print(f"곡정보 저장/이동 실패: {e}")


def _parse_excel_duration_to_hms(value):
    if value is None:
        return None
    try:
        # datetime.time
        if isinstance(value, dt.time):
            return f"{value.hour:02d}", f"{value.minute:02d}", f"{value.second:02d}"
        # datetime.timedelta
        if isinstance(value, dt.timedelta):
            total = int(value.total_seconds())
            hh = total // 3600
            mm = (total % 3600) // 60
            ss = total % 60
            return f"{hh:02d}", f"{mm:02d}", f"{ss:02d}"
        # excel serial number (float or int)
        if isinstance(value, (int, float)):
            # 1 day == 86400 seconds
            total = int(round((float(value) % 1) * 86400))
            hh = total // 3600
            mm = (total % 3600) // 60
            ss = total % 60
            return f"{hh:02d}", f"{mm:02d}", f"{ss:02d}"
        # string like HH:MM:SS or MM:SS
        s = str(value).strip()
        if not s:
            return None
        parts = s.split(":")
        if len(parts) == 3:
            hh, mm, ss = parts
        elif len(parts) == 2:
            hh, mm, ss = "00", parts[0], parts[1]
        else:
            return None
        hh = f"{int(hh):02d}"
        mm = f"{int(mm):02d}"
        ss = f"{int(float(ss)):02d}"
        return hh, mm, ss
    except Exception:
        return None


def _fill_durations_from_excel(driver, excel_path: str) -> None:
    # 트랙 목록 로드 (곡정보 탭 보이게)
    _ensure_tracks_tab(driver)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "track-list")))
    rows = driver.find_elements(By.CSS_SELECTOR, "#track-list tbody tr")
    if not rows:
        print("수록곡 목록을 찾지 못했습니다.")
        return

    # 엑셀 열기 (첫 시트, N열=14번째, 18행부터)
    wb = load_workbook(excel_path, data_only=True)
    ws = wb.active

    updated = 0
    for idx, row_el in enumerate(rows, start=0):
        excel_row = 18 + idx
        cell = ws.cell(row=excel_row, column=14)  # N열
        hms = _parse_excel_duration_to_hms(cell.value)
        if not hms:
            continue

        hh, mm, ss = hms
        try:
            # 해당 행의 곡 링크 클릭하여 폼 로드
            link = WebDriverWait(row_el, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "td[for='displayTrackTitle'] a.show-track-btn"))
            )
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", link)
            link.click()
            # 해당 행의 import_seq와 폼의 #importSeq 동기화 대기
            row_seq = row_el.get_attribute("data-import_seq") or ""
            WebDriverWait(driver, 10).until(
                lambda d: d.find_element(By.ID, "importSeq").get_attribute("value") == row_seq
            )

            # 입력칸 찾기
            hh_input = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "duration_hh")))
            mm_input = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "duration_mm")))
            ss_input = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "duration_ss")))
            hidden_duration = driver.find_element(By.ID, "duration")

            # 실제 키 입력으로 채우기 + 이벤트 트리거
            def type_val(el, val):
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                try:
                    el.click()
                except Exception:
                    pass
                try:
                    el.clear()
                except Exception:
                    # clear가 안되면 전체 선택 후 삭제
                    el.send_keys(Keys.COMMAND, 'a')
                    el.send_keys(Keys.BACK_SPACE)
                el.send_keys(val)
                # keyup/change/input 이벤트 모두 발생
                driver.execute_script(
                    "['keyup','change','input'].forEach(e=>arguments[0].dispatchEvent(new Event(e,{bubbles:true})));",
                    el,
                )

            type_val(hh_input, hh)
            type_val(mm_input, mm)
            type_val(ss_input, ss)

            # hidden #duration도 보정
            if not (hidden_duration.get_attribute("value") or "").strip():
                driver.execute_script(
                    "arguments[0].value = arguments[1]; ['change','input'].forEach(e=>arguments[0].dispatchEvent(new Event(e,{bubbles:true})));",
                    hidden_duration,
                    f"{hh}:{mm}:{ss}",
                )

            updated += 1
        except Exception as e:
            print(f"{idx+1}번째 트랙 재생시간 입력 실패: {e}")

    print(f"재생시간 입력 완료: {updated}개 트랙")


def _ensure_tracks_tab(driver) -> None:
    try:
        # 이미 트랙 탭이 활성화되어 있으면 통과
        meta_tracks = driver.find_element(By.ID, "meta-tracks")
        if "show" in meta_tracks.get_attribute("class") and "active" in meta_tracks.get_attribute("class"):
            return
    except Exception:
        pass

    # 상단 탭에서 곡정보 탭 클릭 시도
    try:
        tab_btn = driver.find_element(By.CSS_SELECTOR, "#metaTab a[data-target='#meta-tracks']")
        tab_btn.click()
    except Exception:
        # 탭 버튼이 없으면 앨범정보의 '곡정보 이동' 버튼으로 이동
        try:
            next_btn = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.ID, "meta-next-album-btn"))
            )
            next_btn.click()
        except Exception:
            pass

    WebDriverWait(driver, 10).until(
        lambda d: "show" in d.find_element(By.ID, "meta-tracks").get_attribute("class") and "active" in d.find_element(By.ID, "meta-tracks").get_attribute("class")
    )


def goto_album_register(driver) -> bool:
    try:
        driver.get(ALBUM_REGISTER_URL)

        WebDriverWait(driver, 10).until(EC.url_contains("/mypage/meta"))
        try:
            WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "form, input, select, textarea"))
            )
        except Exception:
            pass

        print("앨범등록 페이지 진입 완료!")

        bulk_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "register-excel-btn"))
        )
        bulk_btn.click()
        excel_card = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "excel-card"))
        )
        WebDriverWait(driver, 10).until(lambda d: 'd-none' not in excel_card.get_attribute('class'))
        print("대량등록(엑셀) 패널 열기 완료!")

        file_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "mims-excel-upload"))
        )

        excel_filename = "박성태 - 기도왕｜MIMS.xlsx"
        excel_path = os.path.abspath(os.path.join(os.getcwd(), excel_filename))
        if not os.path.exists(excel_path):
            raise FileNotFoundError(f"엑셀 파일을 찾을 수 없습니다: {excel_path}")

        driver.execute_script(
            "arguments[0].classList.remove('d-none'); arguments[0].style.display='block';",
            file_input,
        )
        file_input.send_keys(excel_path)
        print(f"파일 선택 완료: {excel_path}")

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "#excel-card tbody tr"))
        )

        upload_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//div[@id='excel-card']//button[.//span[contains(@class,'fa-upload')]]"))
        )
        upload_btn.click()
        print("업로드 버튼 클릭 완료. 업로드 진행 대기...")

        try:
            WebDriverWait(driver, 5).until(EC.alert_is_present())
            alert = driver.switch_to.alert
            print(f"업로드 경고창 감지: {alert.text}")
            alert.accept()
            print("경고창 확인(accept) 완료")
        except TimeoutException:
            pass

        try:
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#excel-card .badge-success"))
            )
            print("업로드 완료 감지!")
        except TimeoutException:
            print("업로드 성공 배지를 확인하지 못했지만 다음 단계로 진행합니다.")

        try:
            search_btn = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "search-btn"))
            )
            search_btn.click()
        except Exception:
            pass

        table_first_album_xpath = "//div[@id='table']//table//tbody/tr[1]/td[3]/a[contains(@class,'go-register')]"
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, table_first_album_xpath))
        )
        last_err = None
        for attempt in range(5):
            try:
                first_album_link = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, table_first_album_xpath))
                )
                first_album_text = first_album_link.text
                first_album_link.click()
                print(f"상세등록 페이지로 이동 중... (앨범명: {first_album_text})")
                break
            except StaleElementReferenceException as e:
                last_err = e
                time.sleep(0.5)
                continue
        else:
            raise last_err if last_err else Exception("첫 행 앨범 링크 클릭 실패")

        WebDriverWait(driver, 20).until(EC.url_contains("/mypage/meta/register/"))
        print("상세등록 페이지 진입 완료!")

        # 곡정보 탭이 보이지 않으면 이동 시도
        try:
            WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "track-list"))
            )
        except TimeoutException:
            _check_required_and_go_next(driver)
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "track-list"))
            )

        # 업로드에 사용한 엑셀에서 재생시간을 읽어 각 트랙 HH/MM/SS 채우기 (화면 이동/저장은 하지 않음)
        try:
            _fill_durations_from_excel(driver, excel_path)
        except Exception as e:
            print(f"재생시간 채우기 실패: {e}")

        return True
    except Exception as e:
        print(f"앨범등록 페이지/업로드/상세 진입 흐름 실패: {e}")
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
