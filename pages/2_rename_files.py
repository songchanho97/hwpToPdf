import streamlit as st
import os


def rename_files_page():
    """파일 이름을 새로운 규칙에 따라 일괄 변경하는 UI와 로직"""

    st.title("📝 파일 이름 일괄 변경 (넘버링)")
    st.write("지정한 폴더의 모든 파일 이름을 '입력한 이름_번호' 형식으로 변경합니다.")
    st.info("예: `보고서_1.pdf`, `보고서_2.hwp`")

    # 1. 이름을 변경할 파일이 있는 폴더 경로를 입력받는다
    folder_path = st.text_input(
        "📂 이름을 변경할 파일이 있는 폴더 경로", key="rename_folder_path"
    )

    if not folder_path:
        st.info("폴더 경로를 입력해주세요.")
        return

    if not os.path.isdir(folder_path):
        st.error("유효한 폴더 경로가 아닙니다. 경로를 다시 확인해주세요.")
        return

    st.write("---")

    # 2. 사용자로부터 변경할 이름을 입력받는다.
    new_base_name = st.text_input("✏️ 새로운 파일 이름 (번호와 확장자 제외)")

    if st.button("✏️ 이름 변경 실행"):
        if not new_base_name:
            st.warning("새로운 파일 이름을 입력해주세요.")
            return

        try:
            # 폴더 내 파일 목록을 가져옴 (하위 폴더 제외)
            file_list = [
                f
                for f in os.listdir(folder_path)
                if os.path.isfile(os.path.join(folder_path, f))
            ]

            if not file_list:
                st.warning("폴더에 변경할 파일이 없습니다.")
                return

            rename_log = []

            # 3. 모든 파일의 이름을 동일하게 변경하고 "_N" 을 파일명 뒤에 붙인다.
            for i, filename in enumerate(file_list):
                # 파일의 원래 확장자 보존
                _, extension = os.path.splitext(filename)

                # 새로운 파일 이름 생성 (새이름_번호.확장자)
                new_filename = f"{new_base_name}_{i+1}{extension}"

                original_file = os.path.join(folder_path, filename)
                new_file = os.path.join(folder_path, new_filename)

                # 혹시 모를 덮어쓰기 방지
                if os.path.exists(new_file):
                    rename_log.append(
                        f"⚠️ 건너뜀: '{new_filename}' 파일이 이미 존재합니다."
                    )
                    continue

                os.rename(original_file, new_file)
                rename_log.append(f"'{filename}' ➡️ '{new_filename}'")

            st.success(f"총 {len(rename_log)}개의 파일 이름 변경을 완료했습니다.")
            # 변경 내역을 박스 안에 표시
            st.code("\n".join(rename_log), language="text")

        except Exception as e:
            st.error(f"이름 변경 중 오류가 발생했습니다: {e}")


# --- 페이지 실행 ---
rename_files_page()
