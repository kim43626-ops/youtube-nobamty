"""Korean fonts supplied by Debian package or Windows installation."""
from pathlib import Path

def pick(paths):
    for path in paths:
        if Path(path).is_file():
            return path
    raise RuntimeError('한글 글꼴이 없습니다. Streamlit Cloud에서 packages.txt의 fonts-nanum 설치를 확인하세요.')

FONT=pick(['/usr/share/fonts/truetype/nanum/NanumGothic.ttf','C:/Windows/Fonts/malgun.ttf'])
BOLD=pick(['/usr/share/fonts/truetype/nanum/NanumGothicBold.ttf','C:/Windows/Fonts/malgunbd.ttf'])
