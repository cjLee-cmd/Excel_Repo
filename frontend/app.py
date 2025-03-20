import streamlit as st
import pandas as pd
import sys
sys.path.append('/workspaces/Excel_Repo/backend')

from read_excel import read_excel
try:
    from openai import OpenAI
except ModuleNotFoundError:
    st.error("openai 모듈이 설치되지 않았습니다. 'pip install openai' 명령어를 사용하여 설치하세요.")

# tabulate 모듈 확인
try:
    import tabulate
except ModuleNotFoundError:
    st.error("tabulate 모듈이 설치되지 않았습니다. 'pip install tabulate' 명령어를 사용하여 설치하세요.")

st.title('엑셀 파일 뷰어')

# API Key 입력 박스 추가
api_key = st.text_input("OpenAI API Key를 입력하세요", type="password")

uploaded_file = st.file_uploader("엑셀 파일을 업로드하세요", type=["xlsx"])

if uploaded_file is not None:
    try:
        df = read_excel(uploaded_file)
        
        # DataFrame의 모든 열을 문자열로 변환하여 Arrow 호환성 문제 해결
        df = df.astype(str)
        
        # 테이블 형식으로 표시
        st.dataframe(df)
        
        if api_key:
            query = st.text_input("엑셀 파일에 대해 질문하세요")
            if query:
                with st.spinner('답변을 생성 중입니다...'):
                    client = OpenAI(api_key=api_key)
                    
                    # 테이블 모듈을 사용하지 않고 데이터 변환
                    # 컬럼 이름 행 추가
                    header = ','.join(df.columns)
                    # 최대 100개 행 데이터 변환
                    rows = df.head(100).apply(lambda row: ','.join(row), axis=1).tolist()
                    df_info = header + '\n' + '\n'.join(rows)
                    
                    if len(df) > 100:
                        df_info += "\n(참고: 전체 데이터 중 일부만 표시됨)"
                    
                    response = client.chat.completions.create(
                        model="gpt-4o",
                        messages=[
                            {"role": "system", "content": "당신은 엑셀 데이터를 분석하고 질문에 명확하게 답변하는 AI 어시스턴트입니다."},
                            {"role": "user", "content": f"아래 엑셀 데이터를 분석하고 질문에 대해 바로 답변해주세요.\n\n데이터:\n{df_info}\n\n질문: {query}"}
                        ],
                        max_tokens=500  # 더 긴 답변 허용
                    )
                    st.markdown("### 답변")
                    st.markdown(response.choices[0].message.content.strip())
    except Exception as e:
        st.error(f"오류 발생: {str(e)}")
