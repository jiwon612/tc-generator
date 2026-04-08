# 내비게이션 기능 TC 자동 생성기
Claude API를 활용해 TC를 자동 생성하고
그룹별 Excel 시트로 분류 저장하는 도구

## 사용 기술
- Python, Anthropic Claude API, openpyxl

## 실행 방법

### 1. 필수 라이브러리 설치
pip install anthropic openpyxl
### 2. API KEY, 프롬프트 입력
```bash
# ================================
# 설정
# ================================
API_KEY = "여기에_API_키_입력"

```
```bash
USER_PROMPT = """
여기에_프롬프트_입력
"""
```
### 3. 실행
```bash
python main.py
```
## 출력 결과
- 그룹별 시트로 분류된 Excel 파일 자동 생성

<p align="center">
    <img width="400" alt="Image" src="https://github.com/user-attachments/assets/4e75bb6a-5799-4542-8bff-b36c2f643bb7" />
    <br /><br />
    Terminal 출력
    <br /><br />
    <img width="800" alt="Image" src="https://github.com/user-attachments/assets/5620084f-d2cf-4cea-a856-5eddfc211945" />
    <br /><br />
    완성된 전체 TC
    <br /><br />
    <img width="600" alt="Image" src="https://github.com/user-attachments/assets/692b1c27-386a-4c2e-bd3d-e97f18780e77" />
    <br /><br />
    그룹별 시트 분류
    <br /><br />
    <img width="600" alt="Image" src="https://github.com/user-attachments/assets/7bd1b0e9-6dd4-42d0-8bf8-71dbee2dfa42" />
    <br /><br />
    그룹별 TC 생성 요약
</p>
