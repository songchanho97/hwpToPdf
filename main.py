import streamlit as st
import os
import re
import win32com.client
import pythoncom

# Streamlit 웹 페이지 설정
st.set_page_config(page_title="HWP to PDF 변환기", layout="centered")

st.title("📄 HWP 한글 문서를 PDF로 변환")
st.write("HWP 파일이 들어있는 폴더를 선택하고, PDF를 저장할 폴더를 선택하세요.")

# 사용자로부터 입력 받을 폴더 경로
get_path = st.text_input("📂 HWP 파일이 들어있는 폴더 경로")
save_path = st.text_input("💾 PDF 파일을 저장할 폴더 경로")

# 변환 실행 버튼
if st.button("📎 변환 시작"):
    if not os.path.isdir(get_path):
        st.error("HWP 파일 경로가 올바르지 않습니다.")
    elif not os.path.isdir(save_path):
        st.error("저장 경로가 올바르지 않습니다.")
    else:
        try:
            pythoncom.CoInitialize()

            # HWP 프로그램을 제어하기 위한 COM 오브젝트 생성
            hwp = win32com.client.gencache.EnsureDispatch("HWPFrame.HwpObject")
            
            # HWP 보안 모듈 등록 (필수)
            hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")

            # 지정된 폴더 내 .hwp 확장자 파일만 필터링
            files = [f for f in os.listdir(get_path) if re.match(r".*\.hwp$", f)]

            # 파일이 없을 경우 메시지 출력
            if not files:
                st.warning("해당 폴더에 .hwp 파일이 존재하지 않습니다.")
            else:
                for file in files:
                    # HWP 파일 열기
                    hwp.Open(os.path.join(get_path, file))

                    # 파일 이름과 확장자 분리
                    pre, _ = os.path.splitext(file)

                    # PDF로 저장하기 위한 설정값 생성
                    hwp.HAction.GetDefault("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)

                    # 저장할 PDF 파일 경로 설정
                    hwp.HParameterSet.HFileOpenSave.filename = os.path.join(save_path, pre + ".pdf")

                    # 파일 포맷을 PDF로 설정
                    hwp.HParameterSet.HFileOpenSave.Format = "PDF"

                    # 설정된 저장 명령 실행
                    hwp.HAction.Execute("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)

                # HWP 종료
                hwp.Quit()
                st.success(f"총 {len(files)}개의 HWP 파일이 PDF로 변환되었습니다.")
            pythoncom.CoUninitialize()  # ✅ COM 해제

        except Exception as e:
            st.error(f"오류 발생: {e}")