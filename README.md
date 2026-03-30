# Paper-to-PPT Agent

본 레포지토리는 논문 PDF 파일을 읽어 양식에 맞는 PPT를 자동으로 작성해주는 프로젝트입니다.

## How to Install

```bash
# 1. 레포지토리 클론
git clone https://github.com/clary97/paper_presentation.git
cd paper_presentation

# 2. 가상환경 생성 및 활성화
conda create -n pptmaker python=3.12 -y
conda activate pptmaker

# 3. 의존성 설치
pip install -r requirements.txt
```

## Essential

### 1. 템플릿 준비

`templates/` 폴더에 PPT 템플릿 양식(`.pptx`)을 넣어주세요.

```
templates/
└── paper_review-format.pptx
```

### 2. Workspace 구성

`workspaces/` 하위에 **작업 날짜**를 폴더명으로 생성하고, 아래와 같이 구성합니다.

```
workspaces/
└── 260329/                 # 작업 날짜 (YYMMDD)
    ├── 논문.pdf            # 발표할 논문 PDF
    ├── assets/             # 논문 내 Figure, Table 이미지
    │   ├── figure1.png
    │   ├── figure2.png
    │   ├── table1.png
    │   └── ...
    └── output/             # 생성된 PPT 및 중간 파일 (자동 생성)
```

### 3. Assets 준비

`assets/` 폴더에 논문의 Figure, Table을 캡처 또는 저장(이미지 파일)하여 넣어주세요.

- 파일명은 논문의 캡션과 동일하게 지정해야 합니다.
  - `Figure 1` → `figure1.png`
  - `Table 3` → `table3.png`
- 지원 포맷: `.png`, `.jpg`, `.jpeg`, `.gif`, `.bmp`

## Usage

`main.py`의 `TARGET_DATE`를 작업 날짜로 설정한 뒤 실행합니다.

```python
# main.py
TARGET_DATE = "260329"  # 작업할 날짜 폴더명
```

```bash
conda activate pptmaker
python main.py
```

생성된 PPT는 `workspaces/{날짜}/output/` 에 저장됩니다.
