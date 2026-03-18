# -*- coding: utf-8 -*-
"""
KICE 기출문제 다운로더 (KICE Down)
한국교육과정평가원 수능/모의평가 기출문제 자동 다운로드 프로그램

Author: KICE Down Project
Version: 1.0.0
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import sys
import re
import io
import zipfile
import time
import random
import requests
from bs4 import BeautifulSoup

# ─────────────────────────────────────────────
# 상수 및 설정
# ─────────────────────────────────────────────
APP_TITLE = "KICE 기출문제 다운로더"
APP_VERSION = "1.0.0"
BASE_URL = "https://www.suneung.re.kr"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
    "Referer": "https://www.suneung.re.kr/main.do?s=suneung"
}

# 게시판 ID
BOARD_SUNEUNG = "1500234"  # 대학수학능력시험
BOARD_MOCK = "1500236"     # 수능 모의평가

# 연도 범위 (학년도 기준)
YEAR_MIN = 2005
YEAR_MAX = time.localtime().tm_year + 1

# 시험 종류
EXAM_TYPES = {
    "대학수학능력시험": {"board": BOARD_SUNEUNG, "month": None},
    "6월 모의평가": {"board": BOARD_MOCK, "month": "6월"},
    "9월 모의평가": {"board": BOARD_MOCK, "month": "9월"},
}

# ─── 연도별 영역명 매핑 ───
# 평가원 사이트는 연도별로 다른 영역명을 사용함
# 2005~2013: 언어, 수리, 외국어
# 2014: 국어, 수학, 영어 (A/B형)
# 2015~현재: 국어, 수학, 영어
AREA_ALIASES = {
    # 현재 영역명 → 연도별 실제 사이트 영역명
    "국어": {
        "default": "국어",
        "old": "언어",        # 2005~2013
        "cutoff": 2014,      # 2014부터 '국어'
    },
    "수학": {
        "default": "수학",
        "old": "수리",        # 2005~2013
        "cutoff": 2014,
    },
    "영어": {
        "default": "영어",
        "old": "외국어",      # 2005~2013 (외국어(영어))
        "cutoff": 2014,
    },
    # 한국사: 2014 이전에는 '사회탐구' 안에 포함
    "한국사": {
        "default": "한국사",
        "old": None,          # 독립 영역 없음 (사회탐구 ZIP안에 있음)
        "cutoff": 2017,       # 2017부터 독립 영역
    },
}

# ─── 과거 교육과정 ZIP 내 과목명 매핑 ───
# ZIP 안의 파일명이 연도에 따라 다름
# 예: 2020 과학탐구: "과탐(지구 과학 I)" / 2026: "지구과학Ⅰ_문제"
# 예: 2010 이전: CP949 인코딩 '물리 I.PDF'
OLD_SUBJECT_FILTERS = {
    # 과학탐구 과거 과목명 (2014~2020 시기: "과탐(과목 I)" 형식 및 평가원 파일명 오기 방어 로직)
    "물리학I": ["물리학I", "물리학i", "물리i", "물리I", "물리 i", "물리 I", "물리학1", "물리학 1", "물리1", "물리 1"],
    "물리학II": ["물리학II", "물리학ii", "물리ii", "물리II", "물리 ii", "물리 II", "물리학2", "물리학 2", "물리2", "물리 2"],
    "화학I": ["화학I", "화학i", "화학 i", "화학 I", "화학1", "화학 1"],
    "화학II": ["화학II", "화학ii", "화학 ii", "화학 II", "화학2", "화학 2"],
    "생명과학I": ["생명과학I", "생명과학i", "생명 과학 i", "생명 과학 I", "생물I", "생물i", "생물 I", "생물 i", "생명과학1", "생명과학 1", "생명 과학 1", "생물1", "생물 1"],
    "생명과학II": ["생명과학II", "생명과학ii", "생명 과학 ii", "생명 과학 II", "생물II", "생물ii", "생물 II", "생물 ii", "생명과학2", "생명과학 2", "생명 과학 2", "생물2", "생물 2"],
    "지구과학I": ["지구과학I", "지구과학i", "지구 과학 i", "지구 과학 I", "지구과학Ⅰ", "지구과학1", "지구과학 1", "지구 과학 1"],
    "지구과학II": ["지구과학II", "지구과학ii", "지구 과학 ii", "지구 과학 II", "지구과학Ⅱ", "지구과학2", "지구과학 2", "지구 과학 2"],
    # 사회탐구 과거 과목명
    "생활과윤리": ["생활과윤리", "생활과 윤리"],
    "윤리와사상": ["윤리와사상", "윤리와 사상", "윤리"],
    "한국지리": ["한국지리", "한국 지리"],
    "세계지리": ["세계지리", "세계 지리"],
    "동아시아사": ["동아시아사", "동아시아 사"],
    "세계사": ["세계사", "세계 사"],
    "경제": ["경제"],
    "정치와법": ["정치와법", "정치와 법", "법과정치", "법과 정치"],
    "사회문화": ["사회문화", "사회 문화", "사회·문화"],
    "한국사": ["한국사", "국사", "한국근현대사", "한국근·현대사"],
}

# 과목 카테고리 및 세부 과목
SUBJECT_CATEGORIES = {
    "공통 (현행)": {
        "국어": {"area": "국어", "is_bundle": False},
        "수학": {"area": "수학", "is_bundle": False},
        "영어": {"area": "영어", "is_bundle": False},
        "한국사": {"area": "한국사", "is_bundle": False},
    },
    "공통 (구 교육과정)": {
        "언어 (~2013)": {"area": "언어", "is_bundle": False},
        "수리 (~2013)": {"area": "수리", "is_bundle": False},
        "외국어/영어 (~2013)": {"area": "외국어", "is_bundle": False},
    },
    "사회탐구": {
        "사회탐구 (전체)": {"area": "사회탐구", "is_bundle": True, "filter": None},
        "생활과 윤리": {"area": "사회탐구", "is_bundle": True, "filter": "생활과윤리"},
        "윤리와 사상": {"area": "사회탐구", "is_bundle": True, "filter": "윤리와사상"},
        "한국지리": {"area": "사회탐구", "is_bundle": True, "filter": "한국지리"},
        "세계지리": {"area": "사회탐구", "is_bundle": True, "filter": "세계지리"},
        "동아시아사": {"area": "사회탐구", "is_bundle": True, "filter": "동아시아사"},
        "세계사": {"area": "사회탐구", "is_bundle": True, "filter": "세계사"},
        "경제": {"area": "사회탐구", "is_bundle": True, "filter": "경제"},
        "정치와 법": {"area": "사회탐구", "is_bundle": True, "filter": "정치와법"},
        "사회·문화": {"area": "사회탐구", "is_bundle": True, "filter": "사회문화"},
    },
    "과학탐구": {
        "과학탐구 (전체)": {"area": "과학탐구", "is_bundle": True, "filter": None},
        "물리학Ⅰ": {"area": "과학탐구", "is_bundle": True, "filter": "물리학I"},
        "물리학Ⅱ": {"area": "과학탐구", "is_bundle": True, "filter": "물리학II"},
        "화학Ⅰ": {"area": "과학탐구", "is_bundle": True, "filter": "화학I"},
        "화학Ⅱ": {"area": "과학탐구", "is_bundle": True, "filter": "화학II"},
        "생명과학Ⅰ": {"area": "과학탐구", "is_bundle": True, "filter": "생명과학I"},
        "생명과학Ⅱ": {"area": "과학탐구", "is_bundle": True, "filter": "생명과학II"},
        "지구과학Ⅰ": {"area": "과학탐구", "is_bundle": True, "filter": "지구과학I"},
        "지구과학Ⅱ": {"area": "과학탐구", "is_bundle": True, "filter": "지구과학II"},
    },
    "직업탐구": {
        "직업탐구 (전체)": {"area": "직업탐구", "is_bundle": True, "filter": None},
        "성공적인 직업생활": {"area": "직업탐구", "is_bundle": True, "filter": "성공적인직업생활"},
        "농업 기초 기술": {"area": "직업탐구", "is_bundle": True, "filter": "농업기초기술"},
        "공업 일반": {"area": "직업탐구", "is_bundle": True, "filter": "공업일반"},
        "상업 경제": {"area": "직업탐구", "is_bundle": True, "filter": "상업경제"},
        "수산·해운 산업 기초": {"area": "직업탐구", "is_bundle": True, "filter": "수산해운산업기초"},
        "인간 발달": {"area": "직업탐구", "is_bundle": True, "filter": "인간발달"},
    },
    "제2외국어/한문": {
        "제2외국어/한문 (전체)": {"area": "제2외국어/한문", "is_bundle": True, "filter": None},
        "독일어Ⅰ": {"area": "제2외국어/한문", "is_bundle": True, "filter": "독일어"},
        "프랑스어Ⅰ": {"area": "제2외국어/한문", "is_bundle": True, "filter": "프랑스어"},
        "스페인어Ⅰ": {"area": "제2외국어/한문", "is_bundle": True, "filter": "스페인어"},
        "중국어Ⅰ": {"area": "제2외국어/한문", "is_bundle": True, "filter": "중국어"},
        "일본어Ⅰ": {"area": "제2외국어/한문", "is_bundle": True, "filter": "일본어"},
        "러시아어Ⅰ": {"area": "제2외국어/한문", "is_bundle": True, "filter": "러시아어"},
        "아랍어Ⅰ": {"area": "제2외국어/한문", "is_bundle": True, "filter": "아랍어"},
        "베트남어Ⅰ": {"area": "제2외국어/한문", "is_bundle": True, "filter": "베트남어"},
        "한문Ⅰ": {"area": "제2외국어/한문", "is_bundle": True, "filter": "한문"},
    },
}

# 기본 파일명 템플릿
DEFAULT_FILENAME_TEMPLATE = "{연도}_{시험}_{과목}_{유형}"

# ─────────────────────────────────────────────
# 스크래핑 엔진
# ─────────────────────────────────────────────
class KICEScraper:
    """평가원 사이트 스크래핑 및 다운로드 엔진"""

    def __init__(self, log_callback=None, progress_callback=None, speed_mode=False):
        self.session = requests.Session()
        self.session.headers.update(HEADERS)
        self.log = log_callback or print
        self.progress = progress_callback or (lambda v, m: None)
        self._cancelled = False
        self.speed_mode = speed_mode

    def cancel(self):
        self._cancelled = True

    def _is_cancelled(self):
        return self._cancelled

    def fetch_posts(self, board_id, year=None, month=None, area=None):
        """게시판에서 해당 조건의 게시물 목록 가져오기 (모든 페이지)"""
        all_posts = []
        page = 1

        while True:
            if self._is_cancelled():
                return all_posts

            params = {
                "boardID": board_id,
                "m": "0403",
                "s": "suneung",
                "page": str(page),
            }
            if year:
                params["C01"] = str(year)
            if board_id == BOARD_MOCK and month:
                params["C02"] = month
            # 영역 필터
            if area:
                area_param_key = "C03" if board_id == BOARD_MOCK else "C02"
                # 제2외국어/한문의 경우 특수 코드 문제 해결
                if area == "제2외국어/한문":
                    # KICE 서버에서 제2외국어/한문의 Value(예: 2제2외국어/한문15 등) 값이 
                    # 연도별로 자주 끊어지거나 불일치하여 "자료 없음" 오류가 발생합니다.
                    # 따라서 파라미터를 아예 보내지 않고(전체 페이지 호출), 클라이언트 단에서 파이썬으로 텍스트 필터링합니다.
                    pass
                else:
                    params[area_param_key] = area

            try:
                resp = self.session.get(
                    f"{BASE_URL}/boardCnts/list.do",
                    params=params,
                    timeout=20
                )
                resp.raise_for_status()
            except Exception as e:
                self.log(f"[오류] 페이지 {page} 로딩 실패: {e}")
                break

            soup = BeautifulSoup(resp.text, "html.parser")
            table = soup.find("table")
            if not table:
                break

            tbody = table.find("tbody")
            rows = tbody.find_all("tr") if tbody else table.find_all("tr")[1:]

            if not rows:
                break

            for row in rows:
                cols = row.find_all("td")
                if len(cols) < 5:
                    continue

                board_seq = cols[0].get_text(strip=True)
                post_year = cols[1].get_text(strip=True)

                if board_id == BOARD_MOCK:
                    post_month = cols[2].get_text(strip=True)
                    post_area = cols[3].get_text(strip=True)
                    file_col = cols[6] if len(cols) > 7 else cols[-1]
                else:
                    post_month = "수능"
                    post_area = cols[2].get_text(strip=True)
                    file_col = cols[-1]

                # 로컬 필터링 (제2외국어/한문의 경우 - 서버 검색 우회 시 필터링 역할)
                if area == "제2외국어/한문":
                    row_text = row.get_text()
                    keywords = ["제2외국어", "한문", "독일어", "프랑스어", "스페인어", 
                                "중국어", "일본어", "러시아어", "아랍어", "베트남어"]
                    # 행 전체 문자열(영역/제목 등)에 위 키워드 중 하나라도 포함되어 있지 않으면 스킵
                    if not any(k in row_text for k in keywords):
                        continue

                # 파일 링크 추출 (onclick에서 fileSeq 해시 추출)
                files = []
                for a_tag in file_col.find_all("a"):
                    onclick = a_tag.get("onclick", "")
                    match = re.search(r"fn_fileDown\('([a-f0-9]+)'\)", onclick)
                    if match:
                        file_seq = match.group(1)
                        # 파일명은 title 또는 img alt에서 추출
                        fname = a_tag.get("title", "")
                        if not fname:
                            img = a_tag.find("img")
                            if img:
                                fname = img.get("alt", "")
                        files.append({
                            "file_seq": file_seq,
                            "filename": fname,
                            "url": f"{BASE_URL}/boardCnts/fileDown.do?fileSeq={file_seq}"
                        })

                all_posts.append({
                    "board_seq": board_seq,
                    "year": post_year,
                    "month": post_month,
                    "area": post_area,
                    "files": files,
                })

            # 다음 페이지 확인
            paging = soup.find_all("a", href=re.compile(r"page=\d+"))
            max_page = page
            for a in paging:
                href = a.get("href", "")
                m = re.search(r"page=(\d+)", href)
                if m:
                    max_page = max(max_page, int(m.group(1)))

            if page >= max_page:
                break
            page += 1

        return all_posts

    def download_file(self, url, save_path):
        """파일 다운로드"""
        try:
            resp = self.session.get(url, timeout=60, stream=True)
            resp.raise_for_status()

            with open(save_path, "wb") as f:
                for chunk in resp.iter_content(chunk_size=8192):
                    if self._is_cancelled():
                        return False
                    f.write(chunk)

            # 안전 모드: 요청 간 랜덤 딜레이 (IP 차단 방지)
            if not self.speed_mode:
                time.sleep(random.uniform(0.5, 1.5))

            return True
        except Exception as e:
            self.log(f"[오류] 다운로드 실패: {e}")
            return False

    def _fix_zip_filename(self, raw_name):
        """ZIP 내 파일명의 인코딩 문제 수정 (CP949/EUC-KR → UTF-8)"""
        # CP949로 인코딩된 바이트를 복원 시도
        try:
            fixed = raw_name.encode('cp437').decode('cp949')
            return fixed
        except (UnicodeDecodeError, UnicodeEncodeError):
            pass
        try:
            fixed = raw_name.encode('cp437').decode('euc-kr')
            return fixed
        except (UnicodeDecodeError, UnicodeEncodeError):
            pass
        return raw_name

    def extract_subject_from_zip(self, zip_path, subject_filter, save_dir,
                                  filename_base_template, year, exam_short,
                                  subj_name, template, area):
        """ZIP에서 특정 과목 파일만 추출 (부모 이름 상속 및 포괄 파일 지원)"""
        extracted = []
        try:
            with zipfile.ZipFile(zip_path, 'r') as zf:
                for member in zf.namelist():
                    # 디렉토리 엔트리 건너뛰기
                    if member.endswith('/') or member.endswith('\\'):
                        continue
                    # 확장자가 없는 항목 건너뛰기
                    ext = os.path.splitext(member)[1]
                    if not ext:
                        continue

                    # 파일명 인코딩 수정 (과거 ZIP의 CP949 문제)
                    display_name = self._fix_zip_filename(member)

                    # 과목 필터링 및 포괄적 파일 우회
                    actual_subj_name = subj_name
                    if subject_filter and not self._match_subject(display_name, subject_filter):
                        if self._is_generic_file(display_name):
                            actual_subj_name = area  # 영역명("과학탐구" 등)으로 포괄 대표 이름 유지
                        else:
                            continue

                    # 파일명에서 유형 판별 (부모 ZIP 파일의 이름으로부터 '정답' 등의 힌트 상속)
                    file_type = self._detect_file_type(display_name, parent_filename=zip_path)

                    if subject_filter:
                        # 템플릿을 사용하여 파일명 생성
                        save_name = self._build_filename(
                            template, year, exam_short,
                            actual_subj_name, file_type, ext
                        )
                    else:
                        # 전체 추출 시 수정된 파일명 사용
                        save_name = os.path.basename(display_name)

                    save_full = os.path.join(save_dir, save_name)

                    # 동일 파일명이 이미 추출되었으면 건너뛰기
                    if save_full in extracted:
                        continue

                    with zf.open(member) as src, open(save_full, "wb") as dst:
                        dst.write(src.read())

                    extracted.append(save_full)
                    self.log(f"  [추출] {os.path.basename(save_full)}")

        except zipfile.BadZipFile:
            self.log(f"[오류] ZIP 파일 손상: {os.path.basename(zip_path)}")
        except Exception as e:
            self.log(f"[오류] ZIP 추출 실패: {e}")

        return extracted

    def _normalize_roman(self, text):
        """로마숫자를 고유 토큰으로 치환 (I/II 부분매칭 방지)"""
        # 전각 로마숫자 → 반각 치환
        text = text.replace('Ⅲ', 'III').replace('ⅲ', 'III')
        text = text.replace('Ⅱ', 'II').replace('ⅱ', 'II')
        text = text.replace('Ⅰ', 'I').replace('ⅰ', 'I')
        # 대소문자 통일 후 치환 (순서 중요: 긴 것부터)
        # 한글 뒤 또는 문자열 시작 뒤에 오는 I/II/III만 치환
        text = re.sub(r'(?<![a-zA-Z#])[Ii]{3}(?![a-zA-Z#])', '#3#', text)
        text = re.sub(r'(?<![a-zA-Z#])[Ii]{2}(?![a-zA-Z#])', '#2#', text)
        text = re.sub(r'(?<![a-zA-Z#])[Ii](?![a-zA-Z#])', '#1#', text)
        # 한글 바로 뒤의 I/i도 치환 (예: "물리i", "화학I", "지구과학i")
        text = re.sub(r'(?<=[가-힣])[Ii]{3}', '#3#', text)
        text = re.sub(r'(?<=[가-힣])[Ii]{2}', '#2#', text)
        text = re.sub(r'(?<=[가-힣])[Ii](?![a-zA-Z#])', '#1#', text)
        return text

    def _match_subject(self, filename, subject_filter):
        """파일명에서 과목 매칭 (다양한 명칭 변형 지원)"""
        # 기본 정규화: 공백/특수문자 제거, 소문자화
        clean_fname = re.sub(r'[\s_\-·()（）]', '', filename).lower()
        clean_filter = re.sub(r'[\s_\-·()（）]', '', subject_filter).lower()

        # 로마숫자를 고유 토큰으로 치환
        clean_fname = self._normalize_roman(clean_fname)
        clean_filter = self._normalize_roman(clean_filter)

        # 1차: 직접 매칭
        if clean_filter in clean_fname:
            return True

        # 2차: 과거 과목명 변형 매칭 (OLD_SUBJECT_FILTERS 사용)
        alt_filters = OLD_SUBJECT_FILTERS.get(subject_filter, [])
        for alt in alt_filters:
            alt_clean = re.sub(r'[\s_\-·()（）]', '', alt).lower()
            alt_clean = self._normalize_roman(alt_clean)
            if alt_clean in clean_fname:
                return True

        return False

    def _detect_file_type(self, filename, parent_filename=""):
        """파일명에서 유형 판별 (문제지/정답), 부모 ZIP 파일명에서 힌트 상속 포함"""
        lower = filename.lower()
        parent_lower = os.path.basename(parent_filename).lower()

        if "정답" in filename or "answer" in lower or "정답" in parent_filename or "answer" in parent_lower:
            return "정답"
        elif "대본" in filename or "대본" in parent_filename:
            return "듣기대본"
        elif "듣기" in filename or "음원" in filename or "듣기" in parent_filename or "음원" in parent_lower:
            return "듣기음원"
        else:
            # "문제" 키워드가 없어도 정답/듣기가 아니면 문제지로 간주
            # (과거 파일: "언어(홀수형).pdf", "과탐(물리 I).pdf" 등)
            return "문제지"

    def _is_generic_file(self, filename):
        """특정 과목명이 없고 영역 전체를 포괄하는 이름인지 판별 (과도한 필터 방지)"""
        clean_fname = re.sub(r'[\s_\-·()（）]', '', filename)
        # 구체적인 세부 과목을 뜻하는 키워드 (I, II 등 기호를 제외한 핵심 단어들)
        specific_keywords = [
            "물리", "화학", "생명", "생물", "지구", 
            "생활", "윤리", "사상", "지리", "역사", "세계사", "동아시아", "정치", "법", "경제", "사회", "문화",
            "농업", "공업", "상업", "수산", "해운", "인간발달", "가사", "직업생활",
            "국사", "근현대사", "언어", "수리", "외국어"
        ]
        
        for kw in specific_keywords:
            if kw in clean_fname:
                return False  # 하나라도 있으면 특정 세부 과목 파일임
                
        # 특정 과목 키워드가 없으면서 포괄적 의미를 띄는 경우
        generic_keywords = ["탐구", "과학", "사회", "직업", "정답", "해설", "전체"]
        for kw in generic_keywords:
            if kw in clean_fname:
                return True
                
        return False

    def _resolve_area(self, area, year):
        """연도에 따라 영역명을 사이트에서 사용하는 실제 이름으로 변환"""
        alias = AREA_ALIASES.get(area)
        if alias:
            if year < alias["cutoff"]:
                old_name = alias.get("old")
                if old_name:
                    return old_name
                else:
                    # 독립 영역이 없는 경우 (예: 한국사 → 사회탐구 안에 포함)
                    return area  # 그대로 반환, 검색 실패 시 로그에 표시
            return alias["default"]
        return area

    def run_download(self, config):
        """
        메인 다운로드 실행
        config = {
            'years': (start, end),
            'exam_types': [시험종류 리스트],
            'subjects': {과목명: info, ...},
            'save_dir': 저장경로,
            'filename_template': 파일명 템플릿
        }
        """
        self._cancelled = False
        start_year, end_year = config['years']
        exam_types = config['exam_types']
        subjects = config['subjects']
        save_dir = config['save_dir']
        template = config.get('filename_template', DEFAULT_FILENAME_TEMPLATE)

        # 총 작업량 추정
        total_tasks = (end_year - start_year + 1) * len(exam_types) * len(subjects)
        completed = 0
        success_count = 0
        fail_count = 0

        self.log(f"{'='*50}")
        self.log(f"다운로드 시작")
        self.log(f"  연도: {start_year} ~ {end_year}")
        self.log(f"  시험: {', '.join(exam_types)}")
        self.log(f"  과목: {', '.join(subjects.keys())}")
        self.log(f"  저장: {save_dir}")
        self.log(f"  파일명: {template}")
        self.log(f"{'='*50}\n")

        os.makedirs(save_dir, exist_ok=True)

        for exam_name in exam_types:
            if self._is_cancelled():
                self.log("\n[취소] 사용자에 의해 중단되었습니다.")
                break

            exam_info = EXAM_TYPES[exam_name]
            board_id = exam_info["board"]
            month_filter = exam_info["month"]

            # 시험 종류에 따른 짧은 이름
            if exam_name == "대학수학능력시험":
                exam_short = "수능"
            elif exam_name == "6월 모의평가":
                exam_short = "6월"
            elif exam_name == "9월 모의평가":
                exam_short = "9월"
            else:
                exam_short = exam_name

            for year in range(start_year, end_year + 1):
                if self._is_cancelled():
                    break

                # 과목별로 영역 그룹핑 (같은 영역은 한 번만 요청)
                area_groups = {}
                for subj_name, subj_info in subjects.items():
                    area = subj_info["area"]
                    if area not in area_groups:
                        area_groups[area] = []
                    area_groups[area].append((subj_name, subj_info))

                for area, subj_list in area_groups.items():
                    if self._is_cancelled():
                        break

                    # 연도에 따른 영역명 변환
                    actual_area = self._resolve_area(area, year)
                    self.log(f"[검색] {year}학년도 {exam_short} - {actual_area}")

                    posts = self.fetch_posts(
                        board_id,
                        year=year,
                        month=month_filter,
                        area=actual_area
                    )

                    # 게시물 없으면 원래 영역명으로 재시도
                    if not posts and actual_area != area:
                        self.log(f"  '{actual_area}'로 검색 실패, '{area}'로 재시도")
                        posts = self.fetch_posts(
                            board_id, year=year,
                            month=month_filter, area=area
                        )

                    if not posts:
                        self.log(f"  자료 없음 (해당 연도/시험에 자료가 없을 수 있음)")
                        for _ in subj_list:
                            completed += 1
                            self.progress(completed / total_tasks * 100,
                                         f"{completed}/{total_tasks}")
                        continue

                    self.log(f"  게시물 {len(posts)}건 발견")

                    # 각 세부 과목 처리
                    for subj_name, subj_info in subj_list:
                        if self._is_cancelled():
                            break

                        is_bundle = subj_info.get("is_bundle", False)
                        subj_filter = subj_info.get("filter", None)

                        for post in posts:
                            if self._is_cancelled():
                                break

                            for file_info in post["files"]:
                                if self._is_cancelled():
                                    break

                                fname = file_info["filename"]
                                furl = file_info["url"]
                                # 세부 과목(subj_filter) 필터 및 포괄적 낱개 파일 처리
                                actual_subj_name = subj_name
                                ext = os.path.splitext(fname)[1].lower()

                                if not is_bundle or (is_bundle and ext != ".zip"):
                                    if subj_filter and not self._match_subject(fname, subj_filter):
                                        if self._is_generic_file(fname):
                                            actual_subj_name = area  # '과학탐구' 등 통짜 이름 유지
                                        else:
                                            continue

                                file_type = self._detect_file_type(fname)

                                # 듣기 관련 파일 건너뛰기 (영어 듣기 음원/대본)
                                if file_type in ("듣기음원", "듣기대본"):
                                    continue

                                # 파일명 생성 (통짜 파일이면 actual_subj_name이 area(예:과학탐구)가 됨)
                                save_name = self._build_filename(
                                    template, year, exam_short,
                                    actual_subj_name, file_type, ext
                                )

                                if is_bundle and ext == ".zip":
                                    # ZIP 파일: 다운로드 → 추출 → 삭제
                                    temp_zip = os.path.join(save_dir, f"_temp_{fname}")
                                    self.log(f"  [다운로드] {fname}")

                                    if self.download_file(furl, temp_zip):
                                        extracted = self.extract_subject_from_zip(
                                            temp_zip, subj_filter, save_dir,
                                            None, year, exam_short,
                                            subj_name, template, area
                                        )

                                        # 임시 ZIP 삭제
                                        try:
                                            os.remove(temp_zip)
                                        except:
                                            pass

                                        if extracted:
                                            success_count += len(extracted)
                                        else:
                                            if subj_filter:
                                                self.log(f"  [경고] ZIP 내 '{subj_name}' 파일 없음")
                                            fail_count += 1
                                    else:
                                        fail_count += 1

                                elif not is_bundle or (is_bundle and ext != ".zip"):
                                    # 일반 PDF 파일 직접 다운로드 (위에서 이미 필터링/재명명이 완료됨)
                                    save_path = os.path.join(save_dir, save_name)
                                    self.log(f"  [다운로드] {fname} → {save_name}")

                                    if self.download_file(furl, save_path):
                                        success_count += 1
                                    else:
                                        fail_count += 1

                        completed += 1
                        self.progress(
                            completed / total_tasks * 100,
                            f"{completed}/{total_tasks}"
                        )

        self.log(f"\n{'='*50}")
        if self._is_cancelled():
            self.log(f"다운로드 중단됨")
        else:
            self.log(f"다운로드 완료!")
        self.log(f"  성공: {success_count}건")
        self.log(f"  실패: {fail_count}건")
        self.log(f"  저장 위치: {save_dir}")
        self.log(f"{'='*50}")

        return not self._is_cancelled()

    def _build_filename(self, template, year, exam, subject, file_type, ext):
        """파일명 템플릿으로 파일명 생성"""
        name = template.format(
            연도=year,
            시험=exam,
            과목=subject,
            유형=file_type
        )
        # 파일명에 사용할 수 없는 문자 제거
        name = re.sub(r'[<>:"/\\|?*]', '_', name)
        name = re.sub(r'_+', '_', name)  # 중복 언더스코어 정리
        return name + ext


# ─────────────────────────────────────────────
# GUI 애플리케이션
# ─────────────────────────────────────────────
class KICEDownApp:
    """KICE 기출문제 다운로더 GUI"""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title(APP_TITLE)
        self.root.geometry("780x705")
        self.root.minsize(700, 650)
        self.root.resizable(True, True)

        # 아이콘 설정 (없으면 무시)
        try:
            import os
            icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icon.ico")
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
            else:
                self.root.iconbitmap(default="")
        except:
            pass

        self.scraper = None
        self.download_thread = None
        self.is_downloading = False

        # 과목 체크박스 변수
        self.subject_vars = {}

        self._build_ui()
        self._center_window()

    def _center_window(self):
        self.root.update_idletasks()
        w = self.root.winfo_width()
        h = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (w // 2)
        y = (self.root.winfo_screenheight() // 2) - (h // 2)
        self.root.geometry(f"+{x}+{y}")

    def _build_ui(self):
        """UI 구축"""
        # 스타일 설정
        style = ttk.Style()
        style.theme_use("clam")

        # 색상 정의 (사진과 같은 확실한 클래식 그레이톤)
        BG = "#E0E0E0"          # 시스템 기본 회색
        CARD_BG = "#E0E0E0"     # 배경과 동일하게 맞춤 (깔끔한 클래식 스타일)
        PRIMARY = "#000000"
        PRIMARY_HOVER = "#333333"
        TEXT = "#000000"
        TEXT_LIGHT = "#404040"
        BORDER = "#9CA3AF"
        SUCCESS = "#059669"

        self.root.configure(bg=BG)

        # 스타일 커스텀
        style.configure("Title.TLabel", font=("맑은 고딕", 16, "bold"),
                        foreground=PRIMARY, background=BG)
        style.configure("Subtitle.TLabel", font=("맑은 고딕", 9),
                        foreground=TEXT_LIGHT, background=BG)
        style.configure("Section.TLabelframe.Label", font=("맑은 고딕", 10, "bold"),
                        foreground=TEXT)
        style.configure("TLabelframe", background=CARD_BG, bordercolor=BORDER)
        style.configure("TLabelframe.Label", background=CARD_BG)
        style.configure("TLabel", background=CARD_BG, foreground=TEXT,
                        font=("맑은 고딕", 9))
        style.configure("TCheckbutton", background=CARD_BG, foreground=TEXT,
                        font=("맑은 고딕", 9))
        style.configure("TButton", font=("맑은 고딕", 9))
        style.configure("TCombobox", font=("맑은 고딕", 9))

        style.configure("Download.TButton",
                        font=("맑은 고딕", 12, "bold"),
                        padding=(20, 10))
        style.configure("Open.TButton",
                        font=("맑은 고딕", 10),
                        padding=(15, 8))

        # 메인 프레임 (스크롤 가능)
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)

        # ── 타이틀
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill=tk.X, pady=(0, 5))
        title_frame.configure(style="TFrame")
        style.configure("TFrame", background=BG)

        ttk.Label(title_frame, text="📥 KICE 기출문제 다운로더",
                  style="Title.TLabel").pack(anchor="w")
        ttk.Label(title_frame, text="한국교육과정평가원 수능 · 모의평가 기출문제 자동 다운로드",
                  style="Subtitle.TLabel").pack(anchor="w")

        # ── 상단 1행 프레임 (연도 & 시험 종류 가로 배치)
        top_row_frame = ttk.Frame(main_frame)
        top_row_frame.pack(fill=tk.X, pady=(0, 5))

        # ── 1. 연도 선택
        year_frame = ttk.LabelFrame(top_row_frame, text=" 연도 범위 ",
                                     style="Section.TLabelframe")
        year_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=False, padx=(0, 5), ipady=2)

        year_inner = ttk.Frame(year_frame)
        year_inner.pack(fill=tk.BOTH, expand=True, padx=10, pady=4)

        ttk.Label(year_inner, text="시작:").grid(row=0, column=0, padx=(0,5))
        self.year_start = ttk.Combobox(
            year_inner,
            values=[str(y) for y in range(YEAR_MAX, YEAR_MIN - 1, -1)],
            width=8, state="readonly"
        )
        self.year_start.set(str(YEAR_MAX - 2))
        self.year_start.grid(row=0, column=1, padx=(0,10))

        ttk.Label(year_inner, text="끝:").grid(row=0, column=2, padx=(0,5))
        self.year_end = ttk.Combobox(
            year_inner,
            values=[str(y) for y in range(YEAR_MAX, YEAR_MIN - 1, -1)],
            width=8, state="readonly"
        )
        self.year_end.set(str(YEAR_MAX - 1))
        self.year_end.grid(row=0, column=3)

        ttk.Label(year_inner, text="학년도",
                  foreground=TEXT_LIGHT).grid(row=0, column=4, padx=(5,0))

        # ── 2. 시험 종류 (체크박스)
        exam_frame = ttk.LabelFrame(top_row_frame, text=" 시험 종류 ",
                                     style="Section.TLabelframe")
        exam_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, ipady=2)

        exam_inner = ttk.Frame(exam_frame)
        exam_inner.pack(fill=tk.BOTH, expand=True, padx=10, pady=4)

        self.exam_vars = {}
        for i, exam_name in enumerate(EXAM_TYPES.keys()):
            var = tk.BooleanVar(value=True)
            self.exam_vars[exam_name] = var
            # 좁은 가로 공간을 위해 패딩 축소
            ttk.Checkbutton(exam_inner, text=exam_name,
                           variable=var).grid(row=0, column=i, padx=(0,10))

        # ── 3. 과목 선택
        subj_frame = ttk.LabelFrame(main_frame, text=" 과목 선택 ",
                                     style="Section.TLabelframe")
        subj_frame.pack(fill=tk.X, pady=(0, 5), ipady=2)

        # 전체선택/해제 버튼
        subj_top = ttk.Frame(subj_frame)
        subj_top.pack(fill=tk.X, padx=15, pady=(4, 2))
        ttk.Button(subj_top, text="전체 선택",
                   command=lambda: self._toggle_all_subjects(True)).pack(side=tk.LEFT, padx=(0,5))
        ttk.Button(subj_top, text="전체 해제",
                   command=lambda: self._toggle_all_subjects(False)).pack(side=tk.LEFT)

        # 카테고리별 과목 표시 (Notebook 탭)
        self.subj_notebook = ttk.Notebook(subj_frame)
        self.subj_notebook.pack(fill=tk.X, padx=15, pady=(2, 5))

        for cat_name, subjects in SUBJECT_CATEGORIES.items():
            tab = ttk.Frame(self.subj_notebook)
            self.subj_notebook.add(tab, text=cat_name)

            row = 0
            col = 0
            max_cols = 3 if cat_name in ("사회탐구", "과학탐구") else 3
            for subj_name in subjects.keys():
                var = tk.BooleanVar(value=False)
                self.subject_vars[subj_name] = var
                cb = ttk.Checkbutton(tab, text=subj_name, variable=var)
                cb.grid(row=row, column=col, sticky="w", padx=10, pady=2)
                col += 1
                if col >= max_cols:
                    col = 0
                    row += 1

        # 공통 과목은 기본 선택 (사용자 요청: 기본 선택 해제 상태로 시작)
        # for subj in ["국어", "수학", "영어", "한국사"]:
        #     if subj in self.subject_vars:
        #         self.subject_vars[subj].set(True)

        # ── 4. 파일명 형식
        name_frame = ttk.LabelFrame(main_frame, text=" 파일명 형식 ",
                                     style="Section.TLabelframe")
        name_frame.pack(fill=tk.X, pady=(0, 5), ipady=2)

        name_inner = ttk.Frame(name_frame)
        name_inner.pack(fill=tk.X, padx=15, pady=4)

        ttk.Label(name_inner,
                  text="사용 가능 변수: {연도} {시험} {과목} {유형}",
                  foreground=TEXT_LIGHT).pack(anchor="w")

        name_entry_frame = ttk.Frame(name_inner)
        name_entry_frame.pack(fill=tk.X, pady=(5,0))

        ttk.Label(name_entry_frame, text="템플릿:").pack(side=tk.LEFT, padx=(0,5))
        self.filename_template = tk.StringVar(value=DEFAULT_FILENAME_TEMPLATE)
        self.template_entry = ttk.Entry(name_entry_frame,
                                         textvariable=self.filename_template,
                                         font=("맑은 고딕", 10))
        self.template_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,10))

        ttk.Button(name_entry_frame, text="기본값",
                   command=lambda: self.filename_template.set(
                       DEFAULT_FILENAME_TEMPLATE
                   )).pack(side=tk.LEFT)

        # 미리보기
        self.preview_label = ttk.Label(name_inner, text="",
                                        foreground=PRIMARY,
                                        font=("맑은 고딕", 9))
        self.preview_label.pack(anchor="w", pady=(5,0))
        self.filename_template.trace_add("write", self._update_preview)
        self._update_preview()

        # ── 5. 저장 경로
        path_frame = ttk.LabelFrame(main_frame, text=" 저장 경로 ",
                                     style="Section.TLabelframe")
        path_frame.pack(fill=tk.X, pady=(0, 5), ipady=2)

        path_inner = ttk.Frame(path_frame)
        path_inner.pack(fill=tk.X, padx=15, pady=4)

        self.save_path = tk.StringVar(value=os.path.join(
            os.path.expanduser("~"), "Desktop", "KICE_기출문제"
        ))
        self.path_entry = ttk.Entry(path_inner,
                                     textvariable=self.save_path,
                                     state="readonly",
                                     font=("맑은 고딕", 9))
        self.path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,10))
        ttk.Button(path_inner, text="폴더 선택",
                   command=self._select_folder).pack(side=tk.LEFT)

        # ── 6. 실행 버튼 및 프로그레스
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill=tk.X, pady=(2, 2))
        action_frame.configure(style="TFrame")

        self.download_btn = ttk.Button(
            action_frame, text="▶  다운로드 시작",
            style="Download.TButton",
            command=self._start_download
        )
        self.download_btn.pack(side=tk.LEFT, padx=(0, 10))

        self.cancel_btn = ttk.Button(
            action_frame, text="■  중지",
            command=self._cancel_download, state="disabled"
        )
        self.cancel_btn.pack(side=tk.LEFT, padx=(0, 10))

        # 속도 우선 토글
        speed_toggle_frame = ttk.Frame(action_frame)
        speed_toggle_frame.pack(side=tk.LEFT, padx=(10, 0))
        self.speed_mode_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(speed_toggle_frame, text="⚡ 속도 우선",
                        variable=self.speed_mode_var).pack(side=tk.LEFT)
        ttk.Label(speed_toggle_frame, text="(OFF: 안전 딜레이 적용)",
                  foreground=TEXT_LIGHT,
                  font=("맑은 고딕", 8)).pack(side=tk.LEFT, padx=(3, 0))

        self.open_folder_btn = ttk.Button(
            action_frame, text="📂  다운로드 폴더 열기",
            style="Open.TButton",
            command=self._open_folder, state="disabled"
        )
        self.open_folder_btn.pack(side=tk.RIGHT)

        # 프로그레스 바
        prog_frame = ttk.Frame(main_frame)
        prog_frame.pack(fill=tk.X, pady=(2, 2))

        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(
            prog_frame, variable=self.progress_var,
            maximum=100, length=400
        )
        self.progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,10))

        self.progress_label = ttk.Label(prog_frame, text="대기 중",
                                         foreground=TEXT_LIGHT)
        self.progress_label.pack(side=tk.LEFT)

        # ── 7. 로그 창
        log_frame = ttk.LabelFrame(main_frame, text=" 진행 로그 ",
                                    style="Section.TLabelframe")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(2, 0))

        self.log_text = tk.Text(
            log_frame, height=13,
            font=("Consolas", 9),
            bg="#1E293B", fg="#E2E8F0",
            insertbackground="#E2E8F0",
            selectbackground="#334155",
            relief="flat", padx=10, pady=8,
            wrap=tk.WORD
        )
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical",
                                  command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y, padx=(0,5), pady=5)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.log_text.configure(state="disabled")

        # 로그 태그 색상
        self.log_text.tag_configure("info", foreground="#E2E8F0")
        self.log_text.tag_configure("success", foreground="#34D399")
        self.log_text.tag_configure("error", foreground="#F87171")
        self.log_text.tag_configure("warn", foreground="#FBBF24")

    def _update_preview(self, *args):
        """파일명 미리보기 업데이트"""
        try:
            template = self.filename_template.get()
            preview = template.format(
                연도="2026", 시험="수능", 과목="국어", 유형="문제지"
            )
            preview = re.sub(r'[<>:"/\\|?*]', '_', preview)
            self.preview_label.configure(text=f"미리보기: {preview}.pdf")
        except:
            self.preview_label.configure(text="미리보기: (형식 오류)")

    def _toggle_all_subjects(self, state):
        for var in self.subject_vars.values():
            var.set(state)

    def _select_folder(self):
        folder = filedialog.askdirectory(
            title="저장 폴더 선택",
            initialdir=self.save_path.get()
        )
        if folder:
            self.save_path.set(folder)

    def _open_folder(self):
        path = self.save_path.get()
        if os.path.isdir(path):
            os.startfile(path)
        else:
            messagebox.showinfo("알림", "폴더가 존재하지 않습니다.")

    def _log(self, message):
        """스레드 안전한 로그 출력"""
        def _append():
            self.log_text.configure(state="normal")
            # 태그 결정
            tag = "info"
            if "[오류]" in message or "[실패]" in message:
                tag = "error"
            elif "[완료]" in message or "[성공]" in message or "완료" in message:
                tag = "success"
            elif "[경고]" in message:
                tag = "warn"

            self.log_text.insert(tk.END, message + "\n", tag)
            self.log_text.see(tk.END)
            self.log_text.configure(state="disabled")

        self.root.after(0, _append)

    def _update_progress(self, value, text):
        """프로그레스 업데이트"""
        def _update():
            self.progress_var.set(value)
            self.progress_label.configure(text=text)
        self.root.after(0, _update)

    def _validate_config(self):
        """입력 검증"""
        # 연도 검증
        try:
            y_start = int(self.year_start.get())
            y_end = int(self.year_end.get())
            if y_start > y_end:
                messagebox.showwarning("입력 오류",
                    "시작 연도가 끝 연도보다 큽니다.\n연도 범위를 확인하세요.")
                return None
        except ValueError:
            messagebox.showwarning("입력 오류", "연도를 선택하세요.")
            return None

        # 시험 종류 검증
        selected_exams = [name for name, var in self.exam_vars.items() if var.get()]
        if not selected_exams:
            messagebox.showwarning("입력 오류", "시험 종류를 하나 이상 선택하세요.")
            return None

        # 과목 검증
        selected_subjects = {}
        for subj_name, var in self.subject_vars.items():
            if var.get():
                # 카테고리에서 info 찾기
                for cat in SUBJECT_CATEGORIES.values():
                    if subj_name in cat:
                        selected_subjects[subj_name] = cat[subj_name]
                        break

        if not selected_subjects:
            messagebox.showwarning("입력 오류", "과목을 하나 이상 선택하세요.")
            return None

        # 저장 경로 검증
        save_dir = self.save_path.get()
        if not save_dir:
            messagebox.showwarning("입력 오류", "저장 폴더를 선택하세요.")
            return None

        # 파일명 템플릿 검증
        template = self.filename_template.get().strip()
        if not template:
            messagebox.showwarning("입력 오류", "파일명 형식을 입력하세요.")
            return None

        return {
            'years': (y_start, y_end),
            'exam_types': selected_exams,
            'subjects': selected_subjects,
            'save_dir': save_dir,
            'filename_template': template,
        }

    def _start_download(self):
        """다운로드 시작"""
        config = self._validate_config()
        if not config:
            return

        # UI 상태 변경
        self.is_downloading = True
        self.download_btn.configure(state="disabled")
        self.cancel_btn.configure(state="normal")
        self.open_folder_btn.configure(state="disabled")
        self.progress_var.set(0)

        # 로그 초기화
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state="disabled")

        # 스크레이퍼 생성 및 스레드 시작
        self.scraper = KICEScraper(
            log_callback=self._log,
            progress_callback=self._update_progress,
            speed_mode=self.speed_mode_var.get()
        )

        def _run():
            try:
                success = self.scraper.run_download(config)
            except Exception as e:
                self._log(f"[오류] 예기치 않은 오류: {e}")
                success = False
            finally:
                self.root.after(0, lambda: self._on_download_complete(success))

        self.download_thread = threading.Thread(target=_run, daemon=True)
        self.download_thread.start()

    def _cancel_download(self):
        """다운로드 중지"""
        if self.scraper:
            self.scraper.cancel()
            self._log("[중지] 다운로드를 중단합니다...")

    def _on_download_complete(self, success):
        """다운로드 완료 콜백"""
        self.is_downloading = False
        self.download_btn.configure(state="normal")
        self.cancel_btn.configure(state="disabled")
        self.open_folder_btn.configure(state="normal")

        if success:
            self.progress_var.set(100)
            self.progress_label.configure(text="완료!")
            messagebox.showinfo("완료", "모든 다운로드가 완료되었습니다!")
        else:
            self.progress_label.configure(text="중단됨")

    def run(self):
        """앱 실행"""
        self.root.mainloop()


# ─────────────────────────────────────────────
# 메인 실행
# ─────────────────────────────────────────────
if __name__ == "__main__":
    app = KICEDownApp()
    app.run()
