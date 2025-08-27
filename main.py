# 필요한 라이브러리들을 가져옵니다.
import streamlit as st
import os
import re
import win32com.client
import pythoncom
import time

# Streamlit 웹 페이지의 기본 설정을 합니다.
st.set_page_config(page_title="HWP to PDF 변환기", layout="centered")

# 애플리케이션의 제목과 설명을 작성합니다.
st.title("📄 HWP 한글 문서를 PDF로 변환")
st.write("HWP 파일이 들어있는 폴더를 선택하고, PDF를 저장할 폴더를 선택하세요.")
st.info("💡 한/글 프로그램이 설치되어 있어야 정상적으로 동작합니다.")

# --- 사용자 입력 섹션 ---

# HWP 파일이 있는 폴더 경로를 사용자에게 입력받습니다.
get_path = st.text_input("📂 HWP 파일이 들어있는 폴더 경로", key="input_hwp_path")
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
        hwp = None  # hwp 객체를 미리 선언
        try:
            # 2. 지정된 폴더에서 '.hwp' 또는 '.hwpx' 확장자를 가진 파일 목록을 가져옵니다. (수정된 부분)
            files = [
                f
                for f in os.listdir(get_path)
                if re.search(r".*\.(hwp|hwpx)$", f, re.I)
            ]

            if not files:
                st.warning("해당 폴더에 .hwp 또는 .hwpx 파일이 존재하지 않습니다.")
            else:
                st.info(
                    f"총 {len(files)}개의 파일을 PDF로 변환합니다. 잠시만 기다려주세요..."
                )
                progress_bar = st.progress(0)

                pythoncom.CoInitialize()
                hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
                # hwp.Visible = False
                hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")

                # 3. 각 파일을 순회하며 변환 작업을 수행합니다.
                for i, file in enumerate(files):
                    hwp_file_path = os.path.join(get_path, file)
                    pre, _ = os.path.splitext(file)
                    pdf_save_path = os.path.join(save_path, pre + ".pdf")

                    try:
                        hwp.Open(hwp_file_path, "HWP", "forceopen:true")

                        # PDF로 저장 (안정적인 다른 방식)
                        hwp.HAction.GetDefault(
                            "FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet
                        )
                        hwp.HParameterSet.HFileOpenSave.filename = pdf_save_path
                        hwp.HParameterSet.HFileOpenSave.Format = "PDF"
                        hwp.HAction.Execute(
                            "FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet
                        )

                        st.write(f"✅ 성공: '{file}' 파일이 PDF로 변환되었습니다.")
                        time.sleep(0.5)  # 안정성을 위해 잠시 대기

                    except Exception as e:
                        st.error(f"❌ 오류: '{file}' 변환 중 오류가 발생했습니다: {e}")

                    progress_bar.progress((i + 1) / len(files))

                st.balloons()
                st.success(f"✨ 모든 파일 변환이 완료되었습니다. (총 {len(files)}개)")

        except Exception as e:
            st.error(f"⚠️ 시스템 오류가 발생했습니다: {e}")
            st.error(
                "한/글 프로그램이 설치되어 있는지, 실행 중인 한/글 창은 없는지 확인해주세요."
            )

        finally:
            # 4. 오류 발생 여부와 상관없이 항상 한/글 프로그램을 종료합니다.
            if hwp:
                hwp.Quit()
            pythoncom.CoUninitialize()
