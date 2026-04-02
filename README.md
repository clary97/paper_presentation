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

`root/`내에 `templates/` 폴더에 PPT 템플릿 양식(`.pptx`)을 넣어주세요.

```
templates/
└── paper_review-format.pptx
```

> **Tip**: 템플릿은 **슬라이드 마스터**를 활용하여 제작하는 것을 권장합니다. 슬라이드 마스터에서 레이아웃별 placeholder(제목, 본문, 발표자명 등)를 정의해두면, 코드가 해당 placeholder를 인식하여 자동으로 내용을 채워줍니다. PowerPoint에서 `보기 → 슬라이드 마스터`로 편집할 수 있습니다.

### 2. Workspace 구성

`root/`내에 `workspaces/` 하위에 폴더를 생성하고, 아래와 같이 구성합니다. 폴더명은 날짜, 논문명 등 자유롭게 지정할 수 있습니다.

```
workspaces/
└── 260329/                 # 폴더명 (날짜, 논문명 등 자유)
    ├── 논문.pdf            # 발표할 논문 PDF
    ├── assets/             # 논문 내 Figure, Table 이미지
    │   ├── screenshot1.png
    │   ├── screenshot2.png
    │   └── ...
    └── output/             # 생성된 PPT 및 중간 파일 (자동 생성)
```

### 3. Assets 준비

`assets/` 폴더에 논문의 Figure, Table을 캡처 또는 저장(이미지 파일)하여 넣어주세요.

- **파일명은 자유롭게 지정 가능합니다.** 스크린샷 기본 이름 그대로 넣어도 됩니다.
- Claude Code가 이미지를 직접 읽고 논문의 어떤 Figure/Table인지 자동으로 판별하여 매칭합니다.
- 기존 방식(`figure1.png`, `table3.png` 등)도 그대로 사용할 수 있습니다.
- 지원 포맷: `.png`, `.jpg`, `.jpeg`, `.gif`, `.bmp`

## Usage

`main.py`의 `WORKSPACE_NAME`을 작업 폴더명으로 설정합니다.

```python
# main.py
WORKSPACE_NAME = "260329"  # 작업 폴더명
```

PPT 생성은 두 가지 방식으로 사용할 수 있습니다.

### 방법 1: Claude Code와 대화하며 생성 (권장)

Claude Code에서 본 프로젝트 폴더를 열고, 논문에 대한 슬라이드 구조를 요청하면 `slide_structure.json`을 자동으로 생성해줍니다. assets 폴더의 이미지를 직접 읽어서 논문 캡션과 매칭하므로, 파일명을 일일이 수정할 필요가 없습니다.

```bash
# slide_structure.json이 output/ 에 생성된 후
conda activate pptmaker
python -c "
import json
from ppt_builder import build_presentation

with open('workspaces/260329/output/slide_structure.json', 'r') as f:
    slide_data = json.load(f)

prs = build_presentation(slide_data, 'templates/paper_review-format.pptx', 'workspaces/260329/assets')
prs.save('workspaces/260329/output/260329_presentation.pptx')
"
```

### 방법 2: API 키로 전체 자동 실행

`ANTHROPIC_API_KEY` 환경변수를 설정하면 PDF 추출부터 PPT 생성까지 한 번에 실행됩니다.

```bash
export ANTHROPIC_API_KEY="your-api-key"
conda activate pptmaker
python main.py
```

> 기존에 생성된 `slide_structure.json`이 있다면, 재사용 여부를 묻는 프롬프트가 나타납니다. API 재호출 없이 PPT 스타일만 수정하고 싶을 때 유용합니다.

생성된 PPT는 `workspaces/{폴더명}/output/` 에 저장됩니다.
