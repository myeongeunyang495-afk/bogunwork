# 보건대행기관 업무실적 결과보고서 생성기

PDF 방문보고서를 업로드하면 의사, 산업위생관리기사, 간호사 방문건수와 상담인원, 측정 결과, 조치사항, 홍보물 배포 내용을 추출해 DOCX 결과보고서를 생성하는 웹앱입니다.

## Netlify 배포 파일

GitHub에는 아래 파일과 폴더를 올리면 됩니다.

- `index.html`
- `netlify.toml`
- `package.json`
- `package-lock.json`이 생기면 함께 업로드
- `netlify/functions/report.js`
- `.gitignore`
- `README.md`

Python 로컬 실행용 파일인 `app.py`, `app_docx.py`, `실행.bat`은 Netlify 배포에는 필요하지 않습니다.

## Netlify 배포 방법

1. GitHub 저장소에 위 파일을 업로드합니다.
2. Netlify에서 `Add new site` → `Import an existing project`를 선택합니다.
3. GitHub 저장소를 연결합니다.
4. Build command는 `npm run build`로 둡니다.
5. Publish directory는 `.` 로 둡니다.
6. Deploy를 누릅니다.

## 사용 방법

1. 배포된 Netlify 주소로 접속합니다.
2. PDF 방문보고서를 선택합니다.
3. `결과보고서 만들기`를 누릅니다.
4. 미리보기 수치를 확인합니다.
5. `보고서 다운로드`로 DOCX 파일을 받습니다.

## 주의사항

- 스캔 이미지 PDF는 텍스트가 없어 OCR 처리가 필요할 수 있습니다.
- Netlify Functions 요청 용량 제한 때문에 매우 큰 PDF나 100개 이상의 PDF를 한 번에 올리면 실패할 수 있습니다. 이 경우 월별 파일을 여러 묶음으로 나누어 처리하세요.
- 실제 기관 서식에 따라 추출 문구가 다르면 `netlify/functions/report.js`의 정규식 규칙을 조정하면 됩니다.
