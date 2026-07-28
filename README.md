# 팀 예산 관리 시스템

월별 예산 내역을 입력하고 팀원·항목별 사용액을 확인하는 Streamlit 앱입니다.

## 로컬 실행

```bash
python -m pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Community Cloud 배포

1. 이 폴더의 파일을 GitHub 저장소 최상위에 올립니다.
2. [Streamlit Community Cloud](https://share.streamlit.io/)에서 저장소를 선택합니다.
3. Main file path를 `app.py`로 지정하고 Deploy를 누릅니다.

> Community Cloud가 재시작되면 앱에 입력한 데이터가 초기화될 수 있습니다.
> 앱의 CSV 또는 Excel 다운로드 기능으로 데이터를 별도 보관하세요.
