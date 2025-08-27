# 필요한 라이브러리들을 가져옵니다.
import streamlit as st
import os
import re
import win32com.client

# import pythoncom
import time
from pyhwpx import Hwp

# Streamlit 웹 페이지의 기본 설정을 합니다.
st.set_page_config(page_title="HWP to PDF 변환기", layout="centered")

# 애플리케이션의 제목과 설명을 작성합니다.
st.title("📄 HWP 한글 문서를 PDF로 변환")
st.write("HWP 파일이 들어있는 상위 폴더를 선택하고, PDF를 저장할 폴더를 선택하세요.")
st.info("💡 선택한 폴더 및 그 안의 모든 하위 폴더에 있는 HWP 파일을 찾아 변환합니다.")
st.warning("💡 한/글 프로그램이 설치되어 있어야 정상적으로 동작합니다.")


# --- 사용자 입력 섹션 ---

# HWP 파일이 있는 폴더 경로를 사용자에게 입력받습니다.
get_path = st.text_input("📂 HWP 파일이 들어있는 상위 폴더 경로", key="input_hwp_path")
# PDF 파일을 저장할 폴더 경로를 사용자에게 입력받습니다.
save_path = st.text_input("💾 PDF 파일을 저장할 폴더 경로", key="input_pdf_path")

# --- 변환 실행 로직 ---

# '변환 시작' 버튼이 클릭되면 변환 작업을 시작합니다.
if st.button("📎 변환 시작"):
    # 1. 입력된 경로가 유효한지 확인합니다.
    if not get_path or not save_path:
        st.warning("⚠️ HWP 폴더와 저장 폴더 경로를 모두 입력해주세요.")
    elif not os.path.isdir(get_path):
        st.error("❌ HWP 파일 경로가 올바르지 않습니다. 폴더 경로를 확인해주세요.")
    elif not os.path.isdir(save_path):
        st.error("❌ 저장 경로가 올바르지 않습니다. 폴더 경로를 확인해주세요.")
    else:
        # 2. 지정된 폴더와 모든 하위 폴더에서 HWP/HWPX 파일을 재귀적으로 찾습니다.
        files_to_convert = []
        for root, _, files in os.walk(get_path):
            for file in files:
                if re.search(r"\.(hwp|hwpx)$", file, re.I):
                    # 원본 파일의 전체 경로
                    hwp_file_path = os.path.join(root, file)

                    # ✨ 변경된 부분: 폴더 구조 복제 없이, 파일명으로만 PDF 저장 경로 생성
                    base_name, _ = os.path.splitext(file)
                    pdf_filename = base_name + ".pdf"
                    pdf_save_path = os.path.join(save_path, pdf_filename)

                    files_to_convert.append((hwp_file_path, pdf_save_path, file))

        if not files_to_convert:
            st.warning(
                "해당 폴더 및 하위 폴더에 .hwp 또는 .hwpx 파일이 존재하지 않습니다."
            )
        else:
            total_files = len(files_to_convert)
            st.info(
                f"총 {total_files}개의 파일을 PDF로 변환합니다. 잠시만 기다려주세요..."
            )
            progress_bar = st.progress(0)

            # 3. 찾은 각 파일을 순회하며 변환 작업을 수행합니다.
            for i, (hwp_file_path, pdf_save_path, short_filename) in enumerate(
                files_to_convert
            ):
                hwp = None
                try:

                    hwp = Hwp()

                    hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")

                    hwp.Open(hwp_file_path, "HWP", "forceopen:true")

                    # PDF로 저장
                    hwp.HAction.GetDefault(
                        "FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet
                    )
                    hwp.HParameterSet.HFileOpenSave.filename = pdf_save_path
                    hwp.HParameterSet.HFileOpenSave.Format = "PDF"
                    hwp.HAction.Execute(
                        "FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet
                    )

                    st.write(
                        f"✅ 성공: '{short_filename}' 파일이 PDF로 변환되었습니다."
                    )

                except Exception as e:
                    st.error(
                        f"❌ 오류: '{short_filename}' 변환 중 오류가 발생했습니다: {e}"
                    )

                finally:
                    if hwp:
                        hwp.Quit()

                time.sleep(0.5)
                progress_bar.progress((i + 1) / total_files)

            st.balloons()
            st.success(f"✨ 모든 파일 변환이 완료되었습니다. (총 {total_files}개)")
