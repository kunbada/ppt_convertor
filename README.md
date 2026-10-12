# PPT 노트 도구

발표자 노트 삽입·추출, 노트 음성 변환(Azure TTS), PDF → PPT 변환을 하는 브라우저 앱입니다.
서버가 없고 모든 처리는 브라우저 안에서 일어납니다. `index.html` 한 파일이 앱 전체입니다.

## 온라인 사용 (GitHub Pages)
저장소 Settings → Pages → Source: `main` / `(root)` → 저장 후 `https://<계정>.github.io/<저장소>/` 접속.

## 음성 변환 키 보관함
03 음성 변환 탭 → "키 보관함". 아이디/비밀번호로 Azure 키·리전·음성을 AES-256-GCM 암호화(PBKDF2 60만회)해
**이 브라우저(localStorage)** 에만 저장합니다. 다른 PC에서는 "파일로 내보내기" → "파일 가져오기".
- 이것은 서버 로그인이 아니라 **암호화**입니다. 비밀번호를 잊으면 복구할 수 없습니다.
- 내보낸 `ppt-vault-*.json` 은 `.gitignore` 로 제외돼 있습니다. 공개 저장소에 올리지 마세요.
- Azure 키 자체를 소스에 넣지 마세요.

## PPT → PDF / 슬라이드 이미지 (로컬, Windows + PowerPoint)
브라우저에서는 PPTX를 정확히 렌더링할 수 없어 로컬 스크립트로 처리합니다.

    pip install pywin32 python-pptx
    python ppt_export.py 발표.pptx pdf      # PDF
    python ppt_export.py 발표.pptx images   # 슬라이드별 PNG
    python ppt_export.py 발표.pptx flat     # 슬라이드가 이미지 한 장씩인 PPTX (노트 유지)
    python ppt_export.py 발표.pptx all
